# AutoOfficeMCP URLACL 注册脚本。
# 背景：server 使用 HttpListener（内核 http.sys）。Windows 要求先为监听 URL 注册
# “URL 保留”（URLACL）：管理员账户隐式拥有权限，非管理员账户必须先注册一次，
# 否则 server 启动时 HttpListener.Start() 报“拒绝访问”（Access Denied，ErrorCode 5）。
# 注册是机器级、一次性的，重启后仍生效；修改 config.json 的端口/绑定地址后需重跑。
#
# 用法（普通账户直接运行即可，会弹 UAC 自提权）：
#   pwsh -NoProfile -ExecutionPolicy Bypass -File register-urlacl.ps1
# 只查看将注册的 URL 与当前保留状态（不修改系统）：
#   pwsh -NoProfile -ExecutionPolicy Bypass -File register-urlacl.ps1 -CheckOnly
# 只授权给指定账户（比 Everyone 更收紧）：
#   pwsh -NoProfile -ExecutionPolicy Bypass -File register-urlacl.ps1 -User "MACHINE\why"
# 撤销注册：
#   netsh http delete urlacl url=http://127.0.0.1:8010/
param(
    [string]$ConfigPath = "",
    [string]$User = "",
    [switch]$CheckOnly
)

$ErrorActionPreference = "Stop"

# ---- 1) 定位 config.json：优先 exe 同目录（运行时真实读取的那份），其次仓库根 ----
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $candidates = @(
        "F:\c#2\AutoOfficeMCP\config.json",
        (Join-Path $PSScriptRoot "config.json")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) { $ConfigPath = $candidate; break }
    }
}
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or -not (Test-Path -LiteralPath $ConfigPath)) {
    throw "找不到 config.json，请用 -ConfigPath 明确指定路径。"
}

# ---- 2) 解析 transport.bind / transport.port，计算待注册 URL（与 HttpTransportHost 的通配转换一致）----
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$bind = [string]$config.transport.bind
if ([string]::IsNullOrWhiteSpace($bind)) { $bind = "127.0.0.1" }
$bind = $bind.Trim()
$port = [int]$config.transport.port
if ($port -le 0) { $port = 8010 }

if ($bind -in @("*", "0.0.0.0", "+", "::")) {
    $urls = @("http://+:$port/")     # 通配保留：任意主机名/网卡
}
else {
    $urls = @("http://${bind}:$port/")
}

Write-Host "config.json : $ConfigPath"
Write-Host "监听地址    : ${bind}:$port"
foreach ($url in $urls) { Write-Host "注册 URL    : $url" }

# ---- 3) 展示当前保留状态（show 不需要管理员）----
$existing = netsh http show urlacl
foreach ($url in $urls) {
    $found = $existing | Select-String -SimpleMatch $url
    if ($found) {
        Write-Host "当前状态    : $url 已注册" -ForegroundColor Green
    }
    else {
        Write-Host "当前状态    : $url 未注册" -ForegroundColor Yellow
    }
}

if ($CheckOnly) { exit 0 }

# ---- 4) 权限检查：非管理员时弹出 UAC 重新运行本脚本 ----
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "需要管理员权限，弹出 UAC 重新运行本脚本..."
    $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if ([string]::IsNullOrWhiteSpace($pwsh)) {
        $pwsh = (Get-Command powershell -ErrorAction SilentlyContinue).Source
    }
    if ([string]::IsNullOrWhiteSpace($pwsh)) { throw "找不到 pwsh/powershell。" }
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath, "-ConfigPath", $ConfigPath)
    if (-not [string]::IsNullOrWhiteSpace($User)) { $arguments += @("-User", $User) }
    $process = Start-Process -FilePath $pwsh -ArgumentList $arguments -Verb RunAs -Wait -PassThru
    exit $process.ExitCode
}

# ---- 5) 注册（幂等；默认 sddl 授予 Everyone(SID S-1-1-0) 监听权限，不受系统语言影响）----
foreach ($url in $urls) {
    $already = $existing | Select-String -SimpleMatch $url
    if ($already) {
        Write-Host "跳过（已注册）：$url"
        continue
    }
    if ([string]::IsNullOrWhiteSpace($User)) {
        netsh http add urlacl url=$url sddl="D:(A;;GX;;;WD)" | Out-Null
    }
    else {
        netsh http add urlacl url=$url user=$User | Out-Null
    }
    if ($LASTEXITCODE -ne 0) {
        throw "netsh http add urlacl 失败：url=$url user=$User（exit=$LASTEXITCODE）"
    }
    Write-Host "已注册：$url" -ForegroundColor Green
}

Write-Host ""
Write-Host "完成。非管理员账户现在可以直接启动 server。"
Write-Host "撤销命令：netsh http delete urlacl url=<上面注册的URL>"
