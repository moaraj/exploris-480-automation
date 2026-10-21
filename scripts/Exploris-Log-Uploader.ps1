<#
Exploris 480 Log Uploader (Direct to Google Drive)

Keeps the instrument, LC and Windows logs on the Exploris PC mirrored to
    gdrive:Exploris480/Logs/<COMPUTERNAME>/

What it collects
  1. Exploris / Tune logs         C:\ProgramData\Thermo\Exploris\Log (and the rest of that tree)
  2. SII for Xcalibur / Chromeleon / LC driver logs  C:\ProgramData\Dionex\Chromeleon\Log
  3. Xcalibur logs                C:\Xcalibur, C:\Thermo
  4. Our own script logs          D:\Scripts\*.txt, *.csv, *.log (watchdog, ledger, backlog)
  5. Windows Application + System events from Thermo / Dionex / Chromeleon /
     Xcalibur / SII sources, plus disk, USB and network errors, exported to text

How it works
  - Every $IntervalSeconds it copies any new or changed log file into a local
    staging folder (D:\Scripts\LogUploader\stage), then runs `rclone copy` to Drive.
  - Files are read with FileShare.ReadWrite, so logs that Thermo software holds open
    still get copied; the originals are never moved, locked or modified.
  - It only uploads. It never deletes anything locally or on Drive.
  - Runs at BelowNormal priority; files over $MaxFileMB MB upload only their last $TailMB MB. It
    stays out of the acquisition's way. It never touches D:\Production Data.

Usage (PowerShell as the instrument user)
  Show what it would collect:  .\Exploris-Log-Uploader.ps1 -Discover
  Run one pass and exit:       .\Exploris-Log-Uploader.ps1 -Once
  Run continuously:            .\Exploris-Log-Uploader.ps1
  Install as a startup task:   .\Exploris-Log-Uploader.ps1 -Install
  Remove the startup task:     .\Exploris-Log-Uploader.ps1 -Uninstall
#>

[CmdletBinding()]
param(
    [switch]$Discover,
    [switch]$Once,
    [switch]$Install,
    [switch]$Uninstall,
    [int]$IntervalSeconds = 60
)

# ------------------------------------------------------------------ settings
$Rclone       = 'D:\Scripts\rclone\rclone.exe'
$RcloneConf   = 'D:\Scripts\rclone\rclone.conf'
$RemoteRoot   = "gdrive:Exploris480/Logs/$env:COMPUTERNAME"

$WorkDir      = 'D:\Scripts\LogUploader'
$StageDir     = Join-Path $WorkDir 'stage'
$SelfLog      = Join-Path $WorkDir 'log_uploader.txt'
$StateFile    = Join-Path $WorkDir 'state.csv'

$MaxFileMB    = 50      # files bigger than this: upload only the last $TailMB as <name>.TAIL<ext>
$TailMB       = 20
$LookbackDays = 14      # only collect files modified in the last N days
$RescanEvery  = 10      # re-list the folders every N cycles (picks up new daily logs)
$EventsEvery  = 5       # export Windows events every N cycles
$EventDays    = 3       # how many days of Windows events per export

# Instrument-PC log folders (copied completely, filtered by extension)
$FullRoots = @(
    'C:\ProgramData\Dionex\Chromeleon\Log',                        # SII / Chromeleon / LC driver + InstrumentServerError
    'C:\ProgramData\Thermo\Exploris\Log',                          # MS-Log, InstrumentTemperature
    'C:\ProgramData\Thermo Scientific\Instruments\Exploris\6.0\System\Programs\dependencies\logs',  # DataSystem
    'C:\ProgramData\Thermo\Xcalibur\Logs',                         # XcalService
    'C:\Program Files\Thermo\Foundation\LOG',                      # Foundation acquisition service
    'C:\ProgramData\Thermo\Foundation\AuditSystemLogs',
    'C:\ProgramData\Thermo\BootServer\Log',
    'C:\ProgramData\Thermo Scientific\Instruments\TNG\MongoDB\log',
    'C:\ProgramData\Thermo Scientific\FreeStyle\Logs',
    'C:\ProgramData\Thermo Scientific\InstConfigLog',
    'C:\ProgramData\Thermo\Exploris\Instrument\Procedure_Results',
    'C:\Xcalibur\system\programs\debug.log',                       # single file, 78 MB -> tail
    'D:\Scripts'                                                     # watchdog, rclone, mirror logs
)

# Extra folders searched for anything that looks like a log (kept small so a rescan is quick)
$SearchRoots = @(
    "$env:LOCALAPPDATA\Thermo",
    "$env:APPDATA\Thermo",
    "$env:LOCALAPPDATA\Dionex"
)

$LogExtensions = @('.log', '.txt', '.csv', '.xml', '.json', '.trc', '.err', '.audit', '.evt', '.html')
$LogWords      = @('log', 'audit', 'trace', 'error', 'alert', 'diagnos', 'event', 'history', 'report')

# Never collect these
$ExcludeLike = @(
    "$WorkDir*",
    'D:\Scripts\rclone\*',
    'D:\Scripts\exploris-worklist-env\*',
    'D:\Scripts\ExplorisWorklists\generated\*',
    '*\site-packages\*', '*\reports\temp\*', '*\Cache\*',
    '*\Installer*', '*\Setup*', '*\Manuals\*', '*\Help\*',
    'D:\Production Data\*', 'D:\Production_Data\*'
)

$EventSources = 'Thermo|Dionex|Chromeleon|Xcalibur|SII|Exploris|Tune|Foundation|Orbitrap|disk|Ntfs|volmgr|USB|e1i|e1d|Tcpip|Netwtw|Kernel-Power'

# ------------------------------------------------------------------ helpers
function Write-Log([string]$msg, [string]$level = 'INFO') {
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $level, $msg
    Write-Host $line
    try {
        if ((Test-Path $SelfLog) -and ((Get-Item $SelfLog).Length -gt 5MB)) {
            Move-Item $SelfLog ($SelfLog -replace '\.txt$', '_prev.txt') -Force
        }
        Add-Content -Path $SelfLog -Value $line -Encoding UTF8
    } catch { }
}

function Test-Excluded([string]$path) {
    foreach ($p in $ExcludeLike) { if ($path -like $p) { return $true } }
    return $false
}

function Get-StagePath([string]$src) {
    # C:\ProgramData\Thermo\Exploris\Log\a.log  ->  stage\C\ProgramData\Thermo\Exploris\Log\a.log
    $drive = $src.Substring(0, 1)
    Join-Path (Join-Path $StageDir $drive) $src.Substring(3)
}

function Copy-Shared([string]$src, [string]$dst) {
    # Reads files that other programs still hold open; writes to temp, then swaps in.
    $dir = Split-Path $dst -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = "$dst.partial"
    $in = $out = $null
    try {
        $in  = [System.IO.File]::Open($src, 'Open', 'Read', 'ReadWrite, Delete')
        $out = [System.IO.File]::Open($tmp, 'Create', 'Write', 'None')
        $in.CopyTo($out)
    } finally {
        if ($out) { $out.Dispose() }
        if ($in)  { $in.Dispose() }
    }
    Move-Item -LiteralPath $tmp -Destination $dst -Force
    (Get-Item -LiteralPath $dst).LastWriteTimeUtc = (Get-Item -LiteralPath $src).LastWriteTimeUtc
}

function Copy-Tail([string]$src, [string]$dst, [long]$bytes) {
    # Big, still-growing logs (e.g. Xcalibur debug.log): keep only the newest part.
    $dir = Split-Path $dst -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = "$dst.partial"
    $in = $out = $null
    try {
        $in  = [System.IO.File]::Open($src, 'Open', 'Read', 'ReadWrite, Delete')
        $start = [Math]::Max(0, $in.Length - $bytes)
        [void]$in.Seek($start, 'Begin')
        $out = [System.IO.File]::Open($tmp, 'Create', 'Write', 'None')
        $in.CopyTo($out)
    } finally {
        if ($out) { $out.Dispose() }
        if ($in)  { $in.Dispose() }
    }
    Move-Item -LiteralPath $tmp -Destination $dst -Force
}

function Find-LogFiles {
    $cutoff = (Get-Date).AddDays(-$LookbackDays)
    $found = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

    foreach ($root in $FullRoots) {
        if (-not (Test-Path $root)) { continue }
        Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $LogExtensions -contains $_.Extension.ToLower() -and
                           $_.LastWriteTime -gt $cutoff -and
                           -not (Test-Excluded $_.FullName) } |
            ForEach-Object { [void]$found.Add($_.FullName) }
    }

    foreach ($root in $SearchRoots) {
        if (-not $root -or -not (Test-Path $root)) { continue }
        Get-ChildItem -Path $root -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {
                $ext = $_.Extension.ToLower()
                if ($_.LastWriteTime -le $cutoff) { return $false }
                if ($LogExtensions -notcontains $ext) { return $false }
                if (Test-Excluded $_.FullName) { return $false }
                if (@('.log', '.trc', '.err', '.audit') -contains $ext) { return $true }
                $p = $_.FullName.ToLower()
                foreach ($w in $LogWords) { if ($p.Contains($w)) { return $true } }
                return $false
            } |
            ForEach-Object { [void]$found.Add($_.FullName) }
    }
    return @($found)
}

function Export-Events {
    $dir = Join-Path $StageDir 'WindowsEvents'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $start = (Get-Date).AddDays(-$EventDays)
    foreach ($logName in 'Application', 'System') {
        try {
            $ev = Get-WinEvent -FilterHashtable @{ LogName = $logName; StartTime = $start } -ErrorAction Stop |
                  Where-Object { $_.ProviderName -match $EventSources -or $_.Level -le 2 }
            $file = Join-Path $dir ('{0}_{1}_last{2}d.txt' -f $logName, $env:COMPUTERNAME, $EventDays)
            $ev | Sort-Object TimeCreated |
                ForEach-Object {
                    '{0:yyyy-MM-dd HH:mm:ss} | {1,-11} | {2} | id {3} | {4}' -f `
                        $_.TimeCreated, $_.LevelDisplayName, $_.ProviderName, $_.Id,
                        (($_.Message -replace '\s+', ' ') -as [string])
                } | Set-Content -Path $file -Encoding UTF8
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found') {
                Write-Log "Event export ($logName) failed: $($_.Exception.Message)" 'WARN'
            }
        }
    }
}

function Load-State {
    $s = @{}
    if (Test-Path $StateFile) {
        Import-Csv $StateFile | ForEach-Object { $s[$_.Path] = "$($_.Length)|$($_.Ticks)" }
    }
    return $s
}

function Save-State($s) {
    $s.GetEnumerator() | ForEach-Object {
        $parts = $_.Value -split '\|'
        [pscustomobject]@{ Path = $_.Key; Length = $parts[0]; Ticks = $parts[1] }
    } | Export-Csv $StateFile -NoTypeInformation -Encoding UTF8
}

function Invoke-Upload {
    $rcArgs = @(
        'copy', $StageDir, $RemoteRoot,
        '--config', $RcloneConf,
        '--update', '--transfers', '4', '--checkers', '8',
        '--exclude', '*.partial',
        '--retries', '3', '--low-level-retries', '5',
        '--stats', '0', '--log-level', 'ERROR'
    )
    $out = & $Rclone @rcArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Log "rclone exit $LASTEXITCODE : $(($out | Out-String).Trim())" 'ERROR'
        return $false
    }
    return $true
}

# ------------------------------------------------------------------ install / uninstall
$TaskName = 'Exploris Log Uploader'
if ($Install) {
    $script = $MyInvocation.MyCommand.Path
    $action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
               -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) `
                -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
        -Description 'Mirrors Exploris/SII/Xcalibur/Windows logs to gdrive:Exploris480/Logs' -Force | Out-Null
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Installed and started scheduled task '$TaskName' (runs at logon, restarts if it stops)."
    return
}
if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    return
}

# ------------------------------------------------------------------ checks
foreach ($d in $WorkDir, $StageDir) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

if ($Discover) {
    $files = Find-LogFiles
    $files | Sort-Object | ForEach-Object {
        $f = Get-Item -LiteralPath $_
        $tag = if ($f.Length -gt ($MaxFileMB * 1MB)) { "  (last $TailMB MB only)" } else { '' }
        '{0:yyyy-MM-dd HH:mm}  {1,12:N0} B  {2}{3}' -f $f.LastWriteTime, $f.Length, $f.FullName, $tag
    }
    Write-Host "`n$($files.Count) log files would be uploaded to $RemoteRoot"
    return
}

if (-not (Test-Path $Rclone))     { Write-Log "rclone not found at $Rclone" 'ERROR'; exit 1 }
if (-not (Test-Path $RcloneConf)) { Write-Log "rclone config not found at $RcloneConf" 'ERROR'; exit 1 }

# one copy at a time
$mutex = New-Object System.Threading.Mutex($false, 'Global\ExplorisLogUploader')
if (-not $mutex.WaitOne(0)) { Write-Log 'Another copy is already running; exiting.' 'WARN'; exit 0 }

try { (Get-Process -Id $PID).PriorityClass = 'BelowNormal' } catch { }

Write-Log "Started. Interval ${IntervalSeconds}s. Target $RemoteRoot"

# ------------------------------------------------------------------ main loop
$state = Load-State
$files = @()
$cycle = 0
try {
    while ($true) {
        try {
            if ($cycle % $RescanEvery -eq 0) {
                $files = Find-LogFiles
                Write-Log "Watching $($files.Count) log files"
            }
            if ($cycle % $EventsEvery -eq 0) { Export-Events }

            $changed = 0
            foreach ($src in $files) {
                try {
                    $fi = Get-Item -LiteralPath $src -ErrorAction Stop
                    $sig = "$($fi.Length)|$($fi.LastWriteTimeUtc.Ticks)"
                    if ($state[$src] -ne $sig) {
                        $dst = Get-StagePath $src
                        if ($fi.Length -gt ($MaxFileMB * 1MB)) {
                            $dst = [IO.Path]::ChangeExtension($dst, '.TAIL' + $fi.Extension)
                            Copy-Tail $src $dst ($TailMB * 1MB)
                        } else {
                            Copy-Shared $src $dst
                        }
                        $state[$src] = $sig
                        $changed++
                    }
                } catch {
                    Write-Log "Copy failed: $src : $($_.Exception.Message)" 'WARN'
                }
            }

            if ($changed -gt 0 -or $cycle % $EventsEvery -eq 0) {
                if (Invoke-Upload) {
                    if ($changed -gt 0) { Write-Log "Uploaded $changed changed file(s)" }
                    Save-State $state
                }
            }
        } catch {
            Write-Log "Cycle error: $($_.Exception.Message)" 'ERROR'
        }

        if ($Once) { break }
        $cycle++
        Start-Sleep -Seconds $IntervalSeconds
    }
} finally {
    $mutex.ReleaseMutex() | Out-Null
    Write-Log 'Stopped.'
}
