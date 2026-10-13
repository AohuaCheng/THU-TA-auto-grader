# THU-TA-auto-grader Windows 安装脚本（PowerShell）
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $Root

Write-Host "== THU-TA-auto-grader 安装 =="
Write-Host "仓库目录: $Root"

$Workspace = Split-Path -Parent $Root
New-Item -ItemType Directory -Force -Path (Join-Path $Workspace "downloads") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $Workspace "templates") | Out-Null

if (-not (Test-Path ".env") -and (Test-Path ".env.example")) {
    Copy-Item ".env.example" ".env"
    Write-Host "已创建 .env，请编辑 URL 与雨课堂 ID"
}
if (-not (Test-Path "credentials.local.env") -and (Test-Path "credentials.local.env.example")) {
    Copy-Item "credentials.local.env.example" "credentials.local.env"
    Write-Host "已创建 credentials.local.env，请填写账号密码"
}

if (Get-Command uv -ErrorAction SilentlyContinue) {
    Write-Host "使用 uv sync 安装依赖…"
    uv sync
} else {
    Write-Host "使用 venv + pip 安装依赖…"
    python -m venv .venv
    & .\.venv\Scripts\python.exe -m pip install -U pip wheel
    & .\.venv\Scripts\pip.exe install -e .
}

Write-Host ""
Write-Host "安装完成。运行: .\scripts\run.ps1 login"
