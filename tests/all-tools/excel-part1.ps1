param(
    [string]$BaseUrl = "http://127.0.0.1:8010",
    [string]$Token = "",
    [int]$Limit = 0
)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "_common.ps1")
try {
    $result = Invoke-AllToolsPart -App "excel" -Part 1 -Parts 2 -BaseUrl $BaseUrl -Token $Token -Limit $Limit
    if ($result.Fail -gt 0) { exit 1 }
    exit 0
}
catch {
    Write-Host "SETUP FAIL: $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}
