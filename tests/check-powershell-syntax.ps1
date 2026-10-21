$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$failed = $false
Get-ChildItem -LiteralPath (Join-Path $root 'scripts') -Filter '*.ps1' -File | ForEach-Object {
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $_.FullName, [ref]$tokens, [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) {
        $failed = $true
        Write-Host "FAIL: $($_.Name)"
        $parseErrors | ForEach-Object { Write-Host $_.ToString() }
    } else {
        Write-Host "PASS: $($_.Name)"
    }
}
if ($failed) { exit 1 }
