# 用法: .\scripts\run.ps1 login
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $Root

$ArgsList = $args

if (Get-Command uv -ErrorAction SilentlyContinue) {
    if (Test-Path "uv.lock") {
        & uv run python main.py @ArgsList
        exit $LASTEXITCODE
    }
}

$Py = Join-Path $Root ".venv\Scripts\python.exe"
if (Test-Path $Py) {
    & $Py main.py @ArgsList
    exit $LASTEXITCODE
}

Write-Error "请先运行 .\scripts\setup.ps1 安装依赖"
exit 1
