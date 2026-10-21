<#
Exploris Raw Data Watchdog (Direct to Google Drive)
Uploads Exploris 480 data straight from D:\ to Google Drive with rclone. No Google Drive
for desktop and NOTHING cached on C:. Keeps the full folder structure and converts every
.raw to profile (non-centroided) mzML next to the .raw on D:\, then uploads that too.

What it does
  - Creates every folder on Drive (including empty ones); mirrors renames as server-side moves
  - Uploads a file only when it is finished: unchanged for $StableSec and not open for writing
  - rclone checks every upload against Drive's MD5 and resumes big uploads after drops
  - msConvert: --mzML --64 --zlib, no peakPicking (profile kept), one file at a time,
    below-normal CPU priority, PAUSED while a sample is acquiring
  - Conversion audit (startup, every reconcile, whenever the queue empties): every .raw is
    checked against D:\Scripts\mzml_ledger.csv and its .mzML on D:\; missing, outdated or
    incomplete conversions are re-queued. Aborted acquisitions are skipped, not retried.
    Anything unconverted is listed in D:\Scripts\unconverted_report.csv
  - Full reconcile at startup and every $ReconcileMin minutes; retries every 5 min if offline
  - Deletions are NOT mirrored unless $MirrorDeletes = $true (then they go to Drive Trash)

One-time setup (everything lives on D:\Scripts)
  1. Download rclone for Windows 64-bit from https://rclone.org/downloads/ and unzip so that
     rclone.exe is at D:\Scripts\rclone\rclone.exe
  2. Create the Google Drive connection (opens a browser to sign in once):
       D:\Scripts\rclone\rclone.exe config --config D:\Scripts\rclone\rclone.conf
         n -> name: gdrive -> storage: drive -> client_id/secret: blank -> scope: 1
         -> service_account_file: blank -> advanced: n -> auto config: y -> Shared Drive: n
  3. Test:  D:\Scripts\rclone\rclone.exe lsd gdrive: --config D:\Scripts\rclone\rclone.conf
  4. Run:   powershell -ExecutionPolicy Bypass -File "D:\Scripts\Exploris Raw Data Watchdog (Direct to Google Drive).ps1"
#>

# ============================ CONFIG ============================
$Source        = "D:\Production Data\Internal_Experiments"   # where Xcalibur writes data
$Remote        = "gdrive:Exploris480/Internal_Experiments"   # Google Drive folder (same as before)
$RcloneExe     = "D:\Scripts\rclone\rclone.exe"
$RcloneConf    = "D:\Scripts\rclone\rclone.conf"
$Include       = @(".raw", ".mzml", ".sld", ".meth", ".csv", ".pdf", ".txt", ".xlsx")  # @() = all files
$Exclude       = @("~*", "*.tmp", "*.partial", "Thumbs.db", "desktop.ini")
$StableSec     = 120     # file must be unchanged this long before upload
$LockCheckDays = 2       # only lock-test files modified within this many days
$ReconcileMin  = 30      # full sweep interval
$LoopSec       = 5       # queue processing interval
$MirrorDeletes = $false  # $true = deleted items are sent to Google Drive Trash
$BwLimit       = "off"   # e.g. "20M" or "08:00,20M 20:00,off" (20 MB/s by day, unlimited at night)
$Transfers     = 1       # parallel uploads; keep 1 on the instrument PC
$PauseConvertDuringAcq = $true    # no msConvert while a sample is acquiring
$PauseUploadDuringAcq  = $false   # $true = also hold uploads until the sequence is idle
$LogFile       = "D:\Scripts\ms_watchdog_log.txt"
$LogMaxMB      = 10
$TransferLog   = "D:\Scripts\rclone_transfer.log"   # live progress of big upload batches
$MatchByExperimentId = $true    # top-level folders with the same experiment ID = the same Drive folder
$ExpIdPattern  = '^(EXP\d{2}-\d{5}-\d{3})'   # e.g. EXP26-00001-002
$FolderMapFile = "D:\Scripts\folder_map.csv"   # ExpId -> Drive folder; edit to force a choice

# ---- msConvert (profile mzML, saved next to the .raw on D:\) ----
$Convert       = $true
$MzmlRoot      = ""      # "" = next to each .raw in $Source; or a separate root folder on D:
$MsConvert     = "C:\Program Files\ProteoWizard\msconvert.exe"
$MsConvertArgs = @("--mzML", "--64", "--zlib")   # no --filter peakPicking => profile data kept
$ConvertSince  = ""      # "" = convert everything; or a date like "2026-10-01"
$ConvertTimeoutMin = 180
$Staging       = "D:\Scripts\mzml_staging"
$LedgerFile    = "D:\Scripts\mzml_ledger.csv"
$ReportFile    = "D:\Scripts\unconverted_report.csv"
$RetryFailedHours = 12
# ================================================================

$Source = $Source.TrimEnd('\', '/')
$Remote = $Remote.TrimEnd('/')
$MzmlRoot = $MzmlRoot.TrimEnd('\', '/')
if (-not $MzmlRoot) { $MzmlRoot = $Source }
$Sep    = [IO.Path]::DirectorySeparatorChar

$mutex = New-Object System.Threading.Mutex($false, "Global\MSRawWatchdog")
if (-not $mutex.WaitOne(0)) { exit 0 }

# ---------- helpers ----------
function Log([string]$msg) {
    try {
        $dir = Split-Path $LogFile
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        if ((Test-Path -LiteralPath $LogFile) -and ((Get-Item -LiteralPath $LogFile).Length -gt $LogMaxMB * 1MB)) {
            Move-Item -LiteralPath $LogFile -Destination "$LogFile.1" -Force
        }
        "$(Get-Date -Format s)  $msg" | Out-File -Append -FilePath $LogFile -Encoding utf8
    } catch {}
}

function Get-Rel([string]$full) { $full.Substring($Source.Length).TrimStart('\', '/') }
# Experiment-ID folder matching: "EXP26-00001-002 Example Run" on Drive and
# "EXP26-00001-002-Example-Run" on D:\ are treated as ONE folder. The existing Drive
# name always wins, so the team's folder is never duplicated.
$RemoteTop = @{}   # experiment ID -> top-level folder name on Drive (this session)
$SavedMap  = @{}   # persisted choices, so the target folder never flips
if (Test-Path -LiteralPath $FolderMapFile) {
    Import-Csv -LiteralPath $FolderMapFile | ForEach-Object { if ($_.ExpId) { $SavedMap[$_.ExpId.ToUpper()] = $_.DriveFolder } }
}
$DupWarned = New-Object 'System.Collections.Generic.HashSet[string]'

function Save-FolderMap {
    $SavedMap.GetEnumerator() | Sort-Object Name |
        ForEach-Object { [pscustomobject]@{ ExpId = $_.Name; DriveFolder = $_.Value } } |
        Export-Csv -LiteralPath $FolderMapFile -NoTypeInformation -Encoding UTF8
}

function Set-MapChoice([string]$id, [string]$name) {
    $RemoteTop[$id] = $name
    if ($SavedMap[$id] -ne $name) { $SavedMap[$id] = $name; Save-FolderMap }
}

function Get-ExpId([string]$name) {
    if ($MatchByExperimentId -and $name -match $ExpIdPattern) { return $Matches[1].ToUpper() }
    return $null
}

function Get-RemoteTop([string]$localTop) {
    $id = Get-ExpId $localTop
    if ($id -and $RemoteTop.ContainsKey($id)) { return $RemoteTop[$id] }
    if ($id) { Set-MapChoice $id $localTop }   # new experiment: first name seen wins from now on
    return $localTop
}

function Get-RemotePath([string]$full) {
    $r = (Get-Rel $full) -replace '\\', '/'
    if (-not $r) { return $Remote }
    $parts = $r.Split('/', 2)
    $top = Get-RemoteTop $parts[0]
    if ($parts.Count -gt 1) { return "$Remote/$top/$($parts[1])" }
    return "$Remote/$top"
}

function Update-RemoteTop {
    # top-level Drive folders, oldest first, so the team's original folder is chosen
    $r = Invoke-Rclone @("lsf", "--dirs-only", "--format", "tp", "--separator", "|", $Remote)
    if (-not $r.Ok) { return }
    $rows = foreach ($line in ($r.Out -split "`r?`n")) {
        if ($line -match '^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)\|(.+)/$') {
            [pscustomobject]@{ Time = [datetime]$Matches[1]; Name = $Matches[2] }
        }
    }
    $RemoteTop.Clear()
    foreach ($g in ($rows | Where-Object { Get-ExpId $_.Name } | Group-Object { Get-ExpId $_.Name })) {
        $id = $g.Name
        $names = @($g.Group | Sort-Object Time | ForEach-Object { $_.Name })
        # 1. saved choice if that folder still exists, 2. the only folder, 3. the oldest
        if ($SavedMap.ContainsKey($id) -and ($names -contains $SavedMap[$id])) { $choice = $SavedMap[$id] }
        else { $choice = $names[0] }
        Set-MapChoice $id $choice
        if ($names.Count -gt 1 -and $DupWarned.Add($id)) {
            Log "DUPLICATE experiment ID $id on Drive: $(($names | ForEach-Object { '"' + $_ + '"' }) -join ', '). Uploading into `"$choice`". Merge the others (see manual)."
        }
    }
    foreach ($d in (Get-ChildItem -LiteralPath $Source -Directory -ErrorAction SilentlyContinue)) {
        $id = Get-ExpId $d.Name
        if ($id -and $RemoteTop.ContainsKey($id) -and $RemoteTop[$id] -ne $d.Name -and -not $script:MapLogged.Contains($d.Name)) {
            Log "MAPPED   D:\...\$($d.Name)  ->  Drive folder `"$($RemoteTop[$id])`" (same experiment ID)"
            [void]$script:MapLogged.Add($d.Name)
        }
    }
}
$MapLogged = New-Object 'System.Collections.Generic.HashSet[string]'


function Test-Wanted([string]$path) {
    $name = [IO.Path]::GetFileName($path)
    foreach ($x in $Exclude) { if ($name -like $x) { return $false } }
    if ($Include.Count -eq 0) { return $true }
    return $Include -contains ([IO.Path]::GetExtension($path).ToLower())
}

function Test-Locked([string]$path) {
    try { $fs = [IO.File]::Open($path, 'Open', 'Read', 'Read'); $fs.Close(); return $false }   # fails only if open for writing
    catch { return $true }
}

function Test-Ready([IO.FileInfo]$fi) {
    if ($fi.LastWriteTime -gt (Get-Date).AddSeconds(-$StableSec)) { return $false }
    if ($fi.LastWriteTime -gt (Get-Date).AddDays(-$LockCheckDays) -and (Test-Locked $fi.FullName)) { return $false }
    return $true
}

function Invoke-Rclone([string[]]$cmdArgs) {
    $common = @("--retries", "5", "--low-level-retries", "20", "--bwlimit", $BwLimit,
                "--drive-chunk-size", "32M", "--transfers", "$Transfers", "-v")
    if ($RcloneConf) { $common += @("--config", $RcloneConf) }
    try   { $out = & $RcloneExe @cmdArgs @common 2>&1 | Out-String; $ok = ($LASTEXITCODE -eq 0) }
    catch { $out = $_.Exception.Message; $ok = $false }
    return [pscustomobject]@{ Ok = $ok; Out = "$out" }
}

if (-not (Test-Path -LiteralPath $RcloneExe)) { Log "rclone not found at $RcloneExe, exiting"; $mutex.ReleaseMutex(); exit 1 }
if (-not (Test-Path -LiteralPath $Source))    { Log "Source $Source not found, exiting";    $mutex.ReleaseMutex(); exit 1 }

try {
    $tag = if ($RcloneConf) { $RcloneConf } else { $Remote }
    Get-CimInstance Win32_Process -Filter "Name='rclone.exe'" -ErrorAction Stop |
        Where-Object { $_.CommandLine -like "*$tag*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; Log "Stopped leftover rclone (PID $($_.ProcessId)) from a previous run" }
} catch {}

$KnownDirs = New-Object 'System.Collections.Generic.HashSet[string]'
function Ensure-RemoteDir([string]$localDir) {
    $rp = Get-RemotePath $localDir
    if ($KnownDirs.Contains($rp)) { return }
    $r = Invoke-Rclone @("mkdir", $rp)
    if ($r.Ok) { [void]$KnownDirs.Add($rp); Log "DIR      $rp" }
    else { Log "MKDIR FAILED $rp : $($r.Out.Trim())" }
}

function Queue-Tree([string]$dir) {
    Ensure-RemoteDir $dir
    Get-ChildItem -LiteralPath $dir -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object { Ensure-RemoteDir $_.FullName }
    Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { Test-Wanted $_.FullName } | ForEach-Object { $Pending[$_.FullName] = Get-Date }
}

# $true = done (uploaded, already identical, or gone); $false = retry later
function Sync-File([string]$src) {
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return $true }
    $fi = Get-Item -LiteralPath $src
    if (-not (Test-Ready $fi)) { return $false }
    if ($PauseUploadDuringAcq -and (Test-Acquiring)) { return $false }
    if ($script:BulkSet -and $script:BulkSet.Contains($fi.FullName)) { return $false }   # the background batch has it
    $r = Invoke-Rclone @("copyto", $src, (Get-RemotePath $src))
    if (-not $r.Ok) { Log "UPLOAD FAILED $(Get-Rel $src) (will retry): $($r.Out.Trim())"; return $false }
    if ($r.Out -match "Copied") { Log ("UPLOADED {0}  ({1:N1} MB)" -f (Get-Rel $src), ($fi.Length / 1MB)) }
    Queue-Convert $src
    return $true
}

function Invoke-Reconcile {
    if ($script:Bulk) {
        Log "RECONCILE skipped: upload batch from $($script:Bulk.Start.ToString('HH:mm')) still running (progress: $TransferLog)"
        Invoke-ConvertAudit
        return $true
    }
    Log "RECONCILE start"
    Invoke-ConvertAudit      # local, runs even if Google is unreachable
    # 1. Folder structure: list remote dirs once, create any that are missing
    $KnownDirs.Clear()
    $mk = Invoke-Rclone @("mkdir", $Remote)
    if (-not $mk.Ok) { Log "Google Drive unreachable, reconcile skipped: $($mk.Out.Trim())"; return $false }
    [void]$KnownDirs.Add($Remote)
    Update-RemoteTop
    $ls = Invoke-Rclone @("lsf", "-R", "--dirs-only", $Remote)
    if ($ls.Ok) {
        foreach ($line in ($ls.Out -split "`r?`n")) {
            if ($line -and $line.EndsWith('/') -and $line -notmatch '^\d{4}/\d\d/\d\d ') {
                [void]$KnownDirs.Add("$Remote/" + $line.TrimEnd('/'))
            }
        }
    }
    Get-ChildItem -LiteralPath $Source -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object { Ensure-RemoteDir $_.FullName }

    # 2. Files: one rclone copy over every stable, unlocked file (rclone skips identical ones)
    $ready = Get-ChildItem -LiteralPath $Source -Recurse -File -ErrorAction SilentlyContinue |
             Where-Object { (Test-Wanted $_.FullName) -and (Test-Ready $_) }
    $notReady = Get-ChildItem -LiteralPath $Source -Recurse -File -ErrorAction SilentlyContinue |
             Where-Object { (Test-Wanted $_.FullName) -and -not (Test-Ready $_) }
    foreach ($f in $notReady) { $Pending[$f.FullName] = Get-Date }

    if ($ready) {
        # one job per top-level folder so each lands in its (experiment-ID matched) Drive folder
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $jobs = New-Object System.Collections.Generic.List[object]
        $byTop = @{}
        $script:BulkSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($f in $ready) {
            [void]$script:BulkSet.Add($f.FullName)
            $rel = (Get-Rel $f.FullName) -replace '\\', '/'
            $parts = $rel.Split('/', 2)
            if ($parts.Count -gt 1) { $key = $parts[0]; $inner = $parts[1] } else { $key = ""; $inner = $rel }
            if (-not $byTop.ContainsKey($key)) { $byTop[$key] = New-Object System.Collections.Generic.List[string] }
            $byTop[$key].Add($inner)
        }
        foreach ($key in $byTop.Keys) {
            $list = [IO.Path]::GetTempFileName()
            [IO.File]::WriteAllLines($list, [string[]]$byTop[$key], $utf8)   # no BOM
            if ($key) { $sd = Join-Path $Source $key; $dd = "$Remote/$(Get-RemoteTop $key)" } else { $sd = $Source; $dd = $Remote }
            $jobs.Add(@{ Src = $sd; Dst = $dd; List = $list; Count = $byTop[$key].Count })
        }
        Start-BulkUpload $jobs $ready.Count $notReady.Count
    } else { Log "RECONCILE nothing ready to upload, $($notReady.Count) in progress" }
    return $true
}


# ---------- acquisition detection ----------
# A .raw file open for writing (by Xcalibur) = a sample is acquiring right now
$AcqState = $false
$AcqChecked = [datetime]::MinValue
function Test-Acquiring {
    if (((Get-Date) - $script:AcqChecked).TotalSeconds -lt 60) { return $script:AcqState }
    $script:AcqChecked = Get-Date
    $since = (Get-Date).AddDays(-1)
    $busy = $false
    foreach ($f in (Get-ChildItem -LiteralPath $Source -Recurse -File -Filter "*.raw" -ErrorAction SilentlyContinue)) {
        if (($f.LastWriteTime -gt $since -or $f.CreationTime -gt $since) -and (Test-Locked $f.FullName)) { $busy = $true; break }
    }
    if ($busy -ne $script:AcqState) {
        if ($busy) { Log "ACQUISITION running: conversions paused$(if ($PauseUploadDuringAcq) { ' and uploads paused' })" }
        else       { Log "ACQUISITION idle: conversions resume" }
    }
    $script:AcqState = $busy
    return $busy
}

# ---------- msConvert: profile mzML ----------
# Ledger of finished conversions (local, so the audit never has to download from Drive)
#   Rel | RawLength | RawTicks | MzmlLength | ConvertedAt
$ConvertQueue = New-Object 'System.Collections.Generic.List[string]'
$ConvertFails = @{}      # src -> @{ Count; Last }
$Conv   = $null
$AuditWhenIdle = $false
$Ledger = @{}

if ($Convert -and -not $MsConvert) {
    $roots = @("$env:ProgramFiles\ProteoWizard", "${env:ProgramFiles(x86)}\ProteoWizard", "$env:LOCALAPPDATA\Apps")
    $MsConvert = Get-ChildItem -Path $roots -Recurse -Filter "msconvert.exe" -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}
if ($Convert -and (-not $MsConvert -or -not (Test-Path -LiteralPath $MsConvert))) {
    Log "msconvert.exe not found, conversion DISABLED (copying still runs). Install ProteoWizard or set `$MsConvert."
    $Convert = $false
}
if ($Convert) {
    Log "msConvert: $MsConvert"
    if (Test-Path -LiteralPath $LedgerFile) {
        Import-Csv -LiteralPath $LedgerFile | ForEach-Object { $Ledger[$_.Rel] = $_ }
        Log "Ledger loaded: $($Ledger.Count) converted file(s) on record"
    }
}

if ($Convert) {
    # A previous watchdog may have been stopped mid-conversion: stop its msconvert and clear staging
    try {
        Get-CimInstance Win32_Process -Filter "Name='msconvert.exe'" -ErrorAction Stop |
            Where-Object { $_.CommandLine -like "*$Staging*" } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; Log "Stopped leftover msconvert (PID $($_.ProcessId))" }
    } catch {}
    if (Test-Path -LiteralPath $Staging) {
        Start-Sleep -Seconds 1
        Get-ChildItem -LiteralPath $Staging -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-MzmlPath([string]$src) {
    Join-Path $MzmlRoot ([IO.Path]::ChangeExtension((Get-Rel $src), ".mzML"))
}

function Test-MzmlComplete([string]$path) {
    # msConvert writes indexed mzML; a complete file ends with </indexedmzML> (or </mzML>)
    $fs = [IO.File]::OpenRead($path)
    try {
        $n = [int][Math]::Min(4096, $fs.Length)
        if ($n -le 0) { return $false }
        [void]$fs.Seek(-$n, [IO.SeekOrigin]::End)
        $buf = New-Object byte[] $n
        [void]$fs.Read($buf, 0, $n)
        return ([Text.Encoding]::UTF8.GetString($buf) -match '</indexedmzML>|</mzML>')
    } finally { $fs.Close() }
}

function Add-LedgerEntry([string]$src, [long]$mzLen) {
    $fi  = Get-Item -LiteralPath $src
    $row = [pscustomobject]@{ Rel = (Get-Rel $src); RawLength = $fi.Length; RawTicks = $fi.LastWriteTime.Ticks
                              MzmlLength = $mzLen; ConvertedAt = (Get-Date -Format s) }
    $Ledger[$row.Rel] = $row
    $row | Export-Csv -LiteralPath $LedgerFile -Append -NoTypeInformation -Encoding UTF8
}

function Add-ConvertFail([string]$src) {
    $n = 1; if ($ConvertFails.ContainsKey($src)) { $n = $ConvertFails[$src].Count + 1 }
    $ConvertFails[$src] = @{ Count = $n; Last = Get-Date }
    return $n
}

# done | needed | failed | aborted | skipped
function Get-ConvertState([string]$src) {
    if (-not $Convert) { return "skipped" }
    if ([IO.Path]::GetExtension($src).ToLower() -ne ".raw") { return "skipped" }
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return "skipped" }
    $fi = Get-Item -LiteralPath $src
    if ($ConvertSince -and $fi.LastWriteTime -lt [datetime]$ConvertSince) { return "skipped" }

    $rel = Get-Rel $src
    $mz  = Get-MzmlPath $src
    $led = $Ledger[$rel]
    $mzi = $null
    if (Test-Path -LiteralPath $mz -PathType Leaf) { $mzi = Get-Item -LiteralPath $mz }

    if ($led -and $mzi -and [long]$led.RawLength -eq $fi.Length -and [long]$led.RawTicks -eq $fi.LastWriteTime.Ticks `
        -and [long]$led.MzmlLength -eq $mzi.Length) { return "done" }
    # mzML made before the ledger existed: adopt it if it is newer than the raw file
    if (-not $led -and $mzi -and $mzi.Length -gt 0 -and $mzi.LastWriteTime -ge $fi.LastWriteTime) {
        Add-LedgerEntry $src $mzi.Length; return "done"
    }
    if ($ConvertFails.ContainsKey($src)) {
        $f = $ConvertFails[$src]
        if ($f.Aborted) {
            # aborted acquisition: skip until the raw file itself changes
            if ($f.RawLength -eq $fi.Length -and $f.RawTicks -eq $fi.LastWriteTime.Ticks) { return "aborted" }
            $ConvertFails.Remove($src)
        } elseif ($f.Count -ge 3) { return "failed" }
    }
    return "needed"
}

function Queue-Convert([string]$src) {
    if ($ConvertQueue.Contains($src) -or ($Conv -and $Conv.Src -eq $src)) { return }
    if ((Get-ConvertState $src) -eq "needed") { $ConvertQueue.Add($src) }
}

# Full check: every .raw in Source must have a complete, current mzML
function Invoke-ConvertAudit {
    if (-not $Convert) { return }
    # give failed files another chance after $RetryFailedHours
    foreach ($k in @($ConvertFails.Keys)) {
        if (-not $ConvertFails[$k].Aborted -and ((Get-Date) - $ConvertFails[$k].Last).TotalHours -ge $RetryFailedHours) { $ConvertFails.Remove($k) }
    }
    $c = @{ total = 0; done = 0; queued = 0; acquiring = 0; failed = 0; aborted = 0; skipped = 0 }
    $report = New-Object System.Collections.Generic.List[object]
    foreach ($fi in (Get-ChildItem -LiteralPath $Source -Recurse -File -Filter "*.raw" -ErrorAction SilentlyContinue)) {
        $c.total++
        $state = Get-ConvertState $fi.FullName
        switch ($state) {
            "done"    { $c.done++ }
            "skipped" { $c.skipped++ }
            "aborted" { $c.aborted++; $report.Add([pscustomobject]@{ Raw = (Get-Rel $fi.FullName); Status = "ABORTED ACQUISITION (raw never finished), skipped"; RawModified = $fi.LastWriteTime }) }
            "failed"  { $c.failed++;  $report.Add([pscustomobject]@{ Raw = (Get-Rel $fi.FullName); Status = "FAILED 3x, retry after $RetryFailedHours h"; RawModified = $fi.LastWriteTime }) }
            "needed"  {
                if (Test-Ready $fi) {
                    $c.queued++; Queue-Convert $fi.FullName
                    $report.Add([pscustomobject]@{ Raw = (Get-Rel $fi.FullName); Status = "queued"; RawModified = $fi.LastWriteTime })
                } else {
                    $c.acquiring++
                    $report.Add([pscustomobject]@{ Raw = (Get-Rel $fi.FullName); Status = "acquiring or locked"; RawModified = $fi.LastWriteTime })
                }
            }
        }
    }
    $report | Export-Csv -LiteralPath $ReportFile -NoTypeInformation -Encoding UTF8 -Force
    if ($report.Count -eq 0) { "All .raw files converted as of $(Get-Date -Format s)" | Set-Content -LiteralPath $ReportFile }
    Log "AUDIT    raw: $($c.total) | converted: $($c.done) | queued: $($c.queued) | acquiring: $($c.acquiring) | failed: $($c.failed) | aborted runs: $($c.aborted) | skipped: $($c.skipped)"
}

function Start-NextConvert {
    while ($ConvertQueue.Count -gt 0) {
        $src = $ConvertQueue[0]; $ConvertQueue.RemoveAt(0)
        if ((Get-ConvertState $src) -ne "needed") { continue }
        $fi = Get-Item -LiteralPath $src
        if (-not (Test-Ready $fi)) { continue }   # picked up again by the next audit
        $stage = Join-Path $Staging ([guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Force -Path $stage | Out-Null
        $out = [IO.Path]::GetFileNameWithoutExtension($src) + ".mzML"
        $argList = @("`"$src`"") + $MsConvertArgs + @("-o", "`"$stage`"", "--outfile", "`"$out`"")
        try {
            $p = Start-Process -FilePath $MsConvert -ArgumentList $argList -PassThru -NoNewWindow `
                    -RedirectStandardOutput (Join-Path $stage "stdout.txt") -RedirectStandardError (Join-Path $stage "stderr.txt")
        } catch {
            $n = Add-ConvertFail $src
            Log "CONVERT FAILED to start (attempt $n/3) $(Get-Rel $src): $($_.Exception.Message)"
            Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
            if ($n -lt 3) { $ConvertQueue.Add($src) }
            return $null
        }
        $null = $p.Handle   # needed so ExitCode is available later in Windows PowerShell 5.1
        try { $p.PriorityClass = [Diagnostics.ProcessPriorityClass]::BelowNormal } catch {}
        Log "CONVERT  start  $(Get-Rel $src)  ($($ConvertQueue.Count) waiting)"
        return @{ Proc = $p; Src = $src; Stage = $stage; Out = $out; Start = Get-Date }
    }
    return $null
}

function Complete-Convert($c) {
    $mz  = Join-Path $c.Stage $c.Out
    $min = ((Get-Date) - $c.Start).TotalMinutes
    $why = ""
    if ($c.Proc.ExitCode -ne 0)                         { $why = "exit code $($c.Proc.ExitCode)" }
    elseif (-not (Test-Path -LiteralPath $mz))          { $why = "no output file" }
    elseif (-not (Test-MzmlComplete $mz))               { $why = "mzML incomplete (no closing tag)" }

    if (-not $why) {
        $final = Get-MzmlPath $c.Src
        $len   = (Get-Item -LiteralPath $mz).Length
        $fd = Split-Path $final
        if (-not (Test-Path -LiteralPath $fd)) { New-Item -ItemType Directory -Force -Path $fd | Out-Null }
        Move-Item -LiteralPath $mz -Destination $final -Force    # same drive: instant, atomic
        if ((Get-Item -LiteralPath $final).Length -ne $len) {
            $why = "move to final folder incomplete"
        } else {
            Add-LedgerEntry $c.Src $len
            $ConvertFails.Remove($c.Src)
            Log ("CONVERTED {0}  ->  {1:N1} MB mzML in {2:N1} min" -f (Get-Rel $c.Src), ($len / 1MB), $min)
        }
    }
    if ($why) {
        $err = ""
        $ef = Join-Path $c.Stage "stderr.txt"
        if (Test-Path -LiteralPath $ef) { $err = ((Get-Content -LiteralPath $ef -Tail 5) -join " | ") }
        if ($err -match "still being acquired") {
            # Thermo header says acquisition never completed (aborted run). Retrying will not help.
            $rfi = Get-Item -LiteralPath $c.Src
            $ConvertFails[$c.Src] = @{ Count = 99; Last = Get-Date; Aborted = $true; RawLength = $rfi.Length; RawTicks = $rfi.LastWriteTime.Ticks }
            Log "SKIPPED  $(Get-Rel $c.Src): aborted acquisition (raw file never finished). Will not retry unless the file changes."
            Remove-Item -LiteralPath $c.Stage -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
        $n = Add-ConvertFail $c.Src
        Log "CONVERT FAILED (attempt $n/3, $why) $(Get-Rel $c.Src): $err"
        if ($n -lt 3) { $ConvertQueue.Add($c.Src) }
    }
    Remove-Item -LiteralPath $c.Stage -Recurse -Force -ErrorAction SilentlyContinue
}

function Step-Converter {
    if (-not $Convert) { return }
    if ($script:Conv) {
        if ($script:Conv.Proc.HasExited) {
            Complete-Convert $script:Conv; $script:Conv = $null
        } elseif (((Get-Date) - $script:Conv.Start).TotalMinutes -gt $ConvertTimeoutMin) {
            try { $script:Conv.Proc.Kill() } catch {}
            Start-Sleep -Seconds 2
            Log "CONVERT TIMEOUT after $ConvertTimeoutMin min: $(Get-Rel $script:Conv.Src)"
            Complete-Convert $script:Conv; $script:Conv = $null
        } else { return }
    }
    if ($ConvertQueue.Count -eq 0 -and $script:AuditWhenIdle) {
        $script:AuditWhenIdle = $false
        Invoke-ConvertAudit   # confirm nothing was missed once the queue empties
    }
    if ($PauseConvertDuringAcq -and (Test-Acquiring)) { return }   # protect the running sample
    $script:Conv = Start-NextConvert
    if ($script:Conv) { $script:AuditWhenIdle = $true }
}


# ---------- background upload batch (startup catch-up and 30-min sweeps) ----------
$BulkSet = $null     # files the background batch is responsible for
$Bulk = $null        # running rclone job
$BulkJobs = $null    # remaining jobs in this batch
$BulkInfo = $null    # batch start time, log offset, error count
function Start-BulkUpload($jobs, [int]$nReady, [int]$nNotReady) {
    if ((Test-Path -LiteralPath $TransferLog) -and ((Get-Item -LiteralPath $TransferLog).Length -gt 20MB)) {
        Move-Item -LiteralPath $TransferLog -Destination "$TransferLog.1" -Force
    }
    "`r`n===== batch started $(Get-Date -Format s): $nReady file(s) in $($jobs.Count) folder(s) =====" |
        Out-File -Append -FilePath $TransferLog -Encoding utf8
    $script:BulkInfo = @{ Start = Get-Date; Offset = (Get-Item -LiteralPath $TransferLog).Length; Failed = 0 }
    $script:BulkJobs = New-Object System.Collections.Generic.Queue[object]
    foreach ($j in $jobs) { $script:BulkJobs.Enqueue($j) }
    Log "RECONCILE checking/uploading $nReady file(s) in the background, $nNotReady still in progress. Live progress: $TransferLog"
    Start-NextBulkJob
}

function Start-NextBulkJob {
    if (-not $script:BulkJobs -or $script:BulkJobs.Count -eq 0) { $script:Bulk = $null; return }
    $j = $script:BulkJobs.Dequeue()
    $a = @("copy", "`"$($j.Src)`"", "`"$($j.Dst)`"", "--files-from-raw", "`"$($j.List)`"",
           "--order-by", "modtime,descending", "--check-first",
           "--transfers", "$Transfers", "--retries", "5", "--low-level-retries", "20",
           "--bwlimit", "`"$BwLimit`"", "--drive-chunk-size", "32M",
           "-v", "--stats", "30s", "--stats-one-line", "--log-file", "`"$TransferLog`"")
    if ($RcloneConf) { $a += @("--config", "`"$RcloneConf`"") }
    $p = Start-Process -FilePath $RcloneExe -ArgumentList $a -PassThru -NoNewWindow
    $null = $p.Handle
    try { $p.PriorityClass = [Diagnostics.ProcessPriorityClass]::BelowNormal } catch {}
    $script:Bulk = @{ Proc = $p; Job = $j; Start = $script:BulkInfo.Start }
}

function Step-BulkUpload {
    if (-not $script:Bulk -or -not $script:Bulk.Proc.HasExited) { return }
    $j = $script:Bulk.Job
    if ($script:Bulk.Proc.ExitCode -ne 0) { $script:BulkInfo.Failed++ }
    Remove-Item -LiteralPath $j.List -Force -ErrorAction SilentlyContinue
    $script:Bulk = $null
    if ($script:BulkJobs.Count -gt 0) { Start-NextBulkJob; return }

    # whole batch finished
    $new = ""
    try {
        $fs = [IO.File]::Open($TransferLog, 'Open', 'Read', 'ReadWrite')
        [void]$fs.Seek($script:BulkInfo.Offset, 'Begin'); $sr = New-Object IO.StreamReader($fs); $new = $sr.ReadToEnd(); $sr.Close()
    } catch {}
    $n   = ([regex]::Matches($new, "Copied \(")).Count
    $min = [math]::Round(((Get-Date) - $script:BulkInfo.Start).TotalMinutes, 1)
    if ($script:BulkInfo.Failed -eq 0) {
        Log "RECONCILE upload batch done: $n file(s) uploaded in $min min"
    } else {
        Log "RECONCILE upload batch done with errors in $($script:BulkInfo.Failed) folder(s) after $min min ($n uploaded); retrying in 5 min. See $TransferLog"
        $script:reconOk = $false; $script:lastRecon = Get-Date
    }
    $script:BulkInfo = $null
    $script:BulkSet = $null
}

# ---------- watcher ----------
$Pending = @{}
$Queue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'

$fsw = New-Object IO.FileSystemWatcher $Source
$fsw.IncludeSubdirectories = $true
$fsw.InternalBufferSize    = 65536
$fsw.NotifyFilter = [IO.NotifyFilters]'FileName, DirectoryName, LastWrite, Size'

$onEvent = {
    $a = $Event.SourceEventArgs
    $old = $null
    if ($a -is [IO.RenamedEventArgs]) { $old = $a.OldFullPath }
    $Event.MessageData.Enqueue([pscustomobject]@{ Type = $a.ChangeType.ToString(); Path = $a.FullPath; Old = $old })
}
$onError = { $Event.MessageData.Enqueue([pscustomobject]@{ Type = 'Overflow'; Path = $null; Old = $null }) }

$subs = @(
    Register-ObjectEvent $fsw Created -Action $onEvent -MessageData $Queue
    Register-ObjectEvent $fsw Changed -Action $onEvent -MessageData $Queue
    Register-ObjectEvent $fsw Renamed -Action $onEvent -MessageData $Queue
    Register-ObjectEvent $fsw Deleted -Action $onEvent -MessageData $Queue
    Register-ObjectEvent $fsw Error   -Action $onError -MessageData $Queue
)
$fsw.EnableRaisingEvents = $true

Log "WATCHDOG started  $Source  ->  $Remote"
$reconOk   = Invoke-Reconcile
$lastRecon = Get-Date

try {
    while ($true) {
        $e = $null
        while ($Queue.TryDequeue([ref]$e)) {
            try {
                switch ($e.Type) {
                    'Overflow' { Log "Watcher buffer overflow, full reconcile"; $reconOk = Invoke-Reconcile; $lastRecon = Get-Date }

                    'Renamed' {
                        $isDir = Test-Path -LiteralPath $e.Path -PathType Container
                        if ($isDir -or (Test-Wanted $e.Path)) {
                            $oldRP = Get-RemotePath $e.Old; $newRP = Get-RemotePath $e.Path
                            if ($oldRP -eq $newRP) { $r = [pscustomobject]@{ Ok = $false } }   # same experiment folder on Drive: nothing to move
                            else { $r = Invoke-Rclone @("moveto", $oldRP, $newRP) }
                            if ($r.Ok) {
                                Log "RENAMED  $(Get-Rel $e.Old)  ->  $(Get-Rel $e.Path)"
                                if ($isDir) {   # refresh known dir cache under the new name
                                    $oldP = Get-RemotePath $e.Old
                                    @($KnownDirs) | Where-Object { $_ -eq $oldP -or $_.StartsWith("$oldP/") } | ForEach-Object { [void]$KnownDirs.Remove($_) }
                                }
                            }
                            if ($isDir) { Queue-Tree $e.Path } else { $Pending[$e.Path] = Get-Date }
                        }
                    }

                    'Deleted' {
                        $Pending.Remove($e.Path)
                        if ($MirrorDeletes) {
                            $rp = Get-RemotePath $e.Path
                            $r = Invoke-Rclone @("deletefile", $rp)
                            if (-not $r.Ok) { $r = Invoke-Rclone @("purge", $rp); [void]$KnownDirs.Remove($rp) }
                            if ($r.Ok) { Log "TRASHED  $(Get-Rel $e.Path)  (Google Drive Trash)" }
                        }
                    }

                    default {   # Created / Changed
                        if (Test-Path -LiteralPath $e.Path -PathType Container) {
                            if ($e.Type -eq 'Created') { Queue-Tree $e.Path } else { Ensure-RemoteDir $e.Path }
                        }
                        elseif (Test-Wanted $e.Path) { $Pending[$e.Path] = Get-Date }
                    }
                }
            } catch { Log "EVENT ERROR $($e.Type) $($e.Path): $($_.Exception.Message)" }
        }

        foreach ($p in @($Pending.Keys)) {
            try { if (Sync-File $p) { $Pending.Remove($p) } }
            catch { Log "UPLOAD ERROR $(Get-Rel $p): $($_.Exception.Message)" }
        }

        try { Step-BulkUpload } catch { Log "BATCH ERROR: $($_.Exception.Message)" }

        # Full sweep on schedule; retry every 5 min if the last one failed (offline)
        $interval = if ($reconOk) { $ReconcileMin } else { 5 }
        if (((Get-Date) - $lastRecon).TotalMinutes -ge $interval) { $reconOk = Invoke-Reconcile; $lastRecon = Get-Date }

        try { Step-Converter } catch { Log "CONVERTER ERROR: $($_.Exception.Message)" }

        Start-Sleep -Seconds $LoopSec
    }
}
finally {
    $fsw.EnableRaisingEvents = $false
    $subs | ForEach-Object { Unregister-Event -SourceIdentifier $_.Name -ErrorAction SilentlyContinue }
    $fsw.Dispose()
    if ($Bulk -and -not $Bulk.Proc.HasExited) { try { $Bulk.Proc.Kill() } catch {} }
    if ($BulkJobs) { foreach ($j in $BulkJobs) { Remove-Item -LiteralPath $j.List -Force -ErrorAction SilentlyContinue } }
    Log "WATCHDOG stopped"
    $mutex.ReleaseMutex()
}
