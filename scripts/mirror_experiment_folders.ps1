$rc    = "D:\Scripts\rclone\rclone.exe"
$cfg   = "D:\Scripts\rclone\rclone.conf"
$src   = "gdrive:Exploris480/Internal_Experiments"
$local = "D:\Production Data\Internal_Experiments"
$log   = "D:\Scripts\mirror_experiment_folders.log"

$names = & $rc lsf $src --dirs-only --config $cfg
if ($LASTEXITCODE -ne 0) { Add-Content $log "$(Get-Date -f s) rclone lsf failed ($LASTEXITCODE)"; exit 1 }

foreach ($n in $names) {
    $n = $n.TrimEnd('/')
    if ($n -notmatch '^EXP\d{2}-\d{5}-\d{3}') { continue }
    $p = Join-Path $local $n
    if (-not (Test-Path -LiteralPath $p)) {
        New-Item -ItemType Directory -Path $p | Out-Null
        Add-Content $log "$(Get-Date -f s) Created $p"
    }
}

& $rc copy $src $local --include "/EXP*/*.xlsx" --drive-skip-gdocs --config $cfg --log-file $log --log-level INFO
