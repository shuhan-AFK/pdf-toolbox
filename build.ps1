$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $projectDir
try {
    python .\make_icon.py
    python -m PyInstaller --noconfirm --clean .\PDF工具箱_白板版.spec
}
finally {
    Pop-Location
}
