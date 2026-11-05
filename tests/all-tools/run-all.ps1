# 全工具冒烟测试并发入口。
# 用法：
#   pwsh -NoProfile -ExecutionPolicy Bypass -File tests\all-tools\run-all.ps1 -BaseUrl http://127.0.0.1:8010
# 说明：
#   - 不启动/停止 server，也不修改 config.json；
#   - Excel/Word/PPT 各拆成两个进程并发执行，Common 单独执行；
#   - 需要先由 prepare-fixtures.ps1 生成夹具；缺失时本脚本会自动生成。
param(
    [string]$BaseUrl = "http://127.0.0.1:8010",
    [string]$Token = "",
    [switch]$SkipPrepare,
    [switch]$Sequential
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "_common.ps1")

$health = Get-ServerHealth -BaseUrl $BaseUrl -Token $Token
Write-Host "SERVER $BaseUrl transport=$($health.transport) workspace=$($health.workspace)"

$fixtureDir = Join-Path $PSScriptRoot "fixtures"
$requiredFixtures = @(
    (Join-Path $fixtureDir "base.xlsx"),
    (Join-Path $fixtureDir "base.docx"),
    (Join-Path $fixtureDir "base.pptx"),
    (Join-Path $fixtureDir "image.png"),
    (Join-Path $fixtureDir "data.csv")
)
$missing = @($requiredFixtures | Where-Object { -not (Test-Path -LiteralPath $_) })
if ($missing.Count -gt 0) {
    if ($SkipPrepare) {
        Write-Host "缺少夹具，且指定了 -SkipPrepare：" -ForegroundColor Red
        $missing | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        exit 2
    }
    Write-Host "生成夹具..." -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot "prepare-fixtures.ps1")
    $missingAfter = @($requiredFixtures | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missingAfter.Count -gt 0) {
        Write-Host "夹具生成失败：" -ForegroundColor Red
        $missingAfter | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        exit 2
    }
}

$resultDir = Join-Path $PSScriptRoot "results"
$logDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
Get-ChildItem -LiteralPath $resultDir -Filter "*.json" -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $logDir -File -ErrorAction SilentlyContinue | Remove-Item -Force

$parts = @(
    @{ Label = "excel-1"; File = "excel-part1.ps1"; App = "excel" },
    @{ Label = "excel-2"; File = "excel-part2.ps1"; App = "excel" },
    @{ Label = "word-1"; File = "word-part1.ps1"; App = "word" },
    @{ Label = "word-2"; File = "word-part2.ps1"; App = "word" },
    @{ Label = "ppt-1"; File = "ppt-part1.ps1"; App = "ppt" },
    @{ Label = "ppt-2"; File = "ppt-part2.ps1"; App = "ppt" },
    @{ Label = "common"; File = "common.ps1"; App = "common" }
)

$pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
if ([string]::IsNullOrWhiteSpace($pwsh)) {
    $pwsh = (Get-Command powershell -ErrorAction SilentlyContinue).Source
}
if ([string]::IsNullOrWhiteSpace($pwsh)) {
    throw "找不到 pwsh/powershell。"
}

$running = @()
foreach ($part in $parts) {
    $stdout = Join-Path $logDir ("{0}.out.log" -f $part.Label)
    $stderr = Join-Path $logDir ("{0}.err.log" -f $part.Label)
    $arguments = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", (Join-Path $PSScriptRoot $part.File),
        "-BaseUrl", $BaseUrl
    )
    if (-not [string]::IsNullOrWhiteSpace($Token)) {
        $arguments += @("-Token", $Token)
    }
    if ($Sequential) {
        Write-Host "RUN $($part.Label) (sequential)" -ForegroundColor Cyan
        & $pwsh @arguments
        $exitCode = $LASTEXITCODE
        $running += [pscustomobject]@{ Label = $part.Label; App = $part.App; ExitCode = $exitCode; Stdout = $stdout; Stderr = $stderr }
    }
    else {
        Write-Host "START $($part.Label)" -ForegroundColor Cyan
        $process = Start-Process -FilePath $pwsh -ArgumentList $arguments -PassThru -WindowStyle Hidden `
            -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $running += [pscustomobject]@{ Label = $part.Label; App = $part.App; Process = $process; ExitCode = $null; Stdout = $stdout; Stderr = $stderr }
    }
}

if (-not $Sequential) {
    foreach ($item in $running) {
        $null = $item.Process.WaitForExit(1800000)
        if (-not $item.Process.HasExited) {
            try { $item.Process.Kill() } catch { }
            $item.ExitCode = 124
        }
        else {
            $item.ExitCode = $item.Process.ExitCode
        }
    }
}

# 汇总报告
$total = 0
$pass = 0
$gated = 0
$fail = 0
$reports = @()
foreach ($item in $running) {
    $reportPath = Join-Path $resultDir ("{0}-part{1}.json" -f $item.App, $(if ($item.Label -like "*-2") { 2 } else { 1 }))
    if ($item.Label -eq "common") {
        $reportPath = Join-Path $resultDir "common-part1.json"
    }
    if (Test-Path -LiteralPath $reportPath) {
        $report = Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $reports += [pscustomobject]@{ Label = $item.Label; Report = $report }
        $total += [int]$report.total
        $pass += [int]$report.pass
        $gated += [int]$report.gated
        $fail += [int]$report.fail
    }
    else {
        Write-Host "MISSING REPORT $($item.Label) exit=$($item.ExitCode)" -ForegroundColor Red
        $fail++
    }
    if ($item.ExitCode -ne 0) {
        Write-Host "PART FAIL $($item.Label) exit=$($item.ExitCode)" -ForegroundColor Red
        if (Test-Path -LiteralPath $item.Stderr) {
            Get-Content -LiteralPath $item.Stderr -Tail 40 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkYellow }
        }
    }
}

# 覆盖率对照：/mcp 全量 tools/list 与各分片总数比较。
$coverageOk = $true
try {
    $sessionId = Initialize-McpSession -BaseUrl $BaseUrl -Token $Token
    $allTools = Get-McpTools -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token
    $expected = @{
        excel = @($allTools | Where-Object { [string]$_.name -like "excel_*" }).Count
        word = @($allTools | Where-Object { [string]$_.name -like "word_*" }).Count
        ppt = @($allTools | Where-Object { [string]$_.name -like "ppt_*" }).Count
        common = @($allTools | Where-Object { [string]$_.name -like "office_*" }).Count
    }
    foreach ($app in @("excel", "word", "ppt", "common")) {
        $actual = 0
        foreach ($item in $reports) {
            if ($item.Report.app -eq $app) { $actual += [int]$item.Report.total }
        }
        if ($actual -ne $expected[$app]) {
            Write-Host "COVERAGE MISMATCH $app actual=$actual expected=$($expected[$app])" -ForegroundColor Red
            $coverageOk = $false
        }
        else {
            Write-Host "COVERAGE $app $actual/$($expected[$app])" -ForegroundColor Green
        }
    }
}
catch {
    Write-Host "覆盖率检查失败：$($_.Exception.Message)" -ForegroundColor Red
    $coverageOk = $false
}

Write-Host ""
Write-Host ("TOTAL total={0} pass={1} gated={2} fail={3}" -f $total, $pass, $gated, $fail) `
    -ForegroundColor $(if ($fail -eq 0 -and $coverageOk) { "Green" } else { "Red" })
Write-Host "GATED 表示参数/安全/状态门禁拒绝，详情见 results/*.json；FAIL 才是需要修复的问题。"

$exit = 0
if ($fail -gt 0 -or -not $coverageOk) { $exit = 1 }
foreach ($item in $running) {
    if ($null -ne $item.ExitCode -and $item.ExitCode -ne 0) { $exit = 1 }
}
exit $exit
