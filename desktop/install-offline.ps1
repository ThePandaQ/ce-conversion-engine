param(
    [string]$InstallDir = "$env:LOCALAPPDATA\CEConversionEngine",
    [switch]$NoDesktopShortcut
)
$ErrorActionPreference = 'Stop'
$src = $PSScriptRoot
$version = '1.0.1'
$verFile = Join-Path $src 'VERSION'
if (Test-Path $verFile) { $version = (Get-Content -LiteralPath $verFile -Raw -Encoding UTF8).Trim() }
Write-Host "=============================================="
Write-Host " CE转换引擎（Conversion Engine）v$version 离线安装"
Write-Host " 作者：温酒会知音    版权所有 © 2026 温酒会知音"
Write-Host "=============================================="
Write-Host "安装目录：$InstallDir"
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
Write-Host "正在复制程序与离线运行环境（约 400MB，请稍候）..."
Copy-Item (Join-Path $src 'app') $InstallDir -Recurse -Force
Copy-Item (Join-Path $src 'python') $InstallDir -Recurse -Force
foreach ($f in @('VERSION','LICENSE','NOTICE.md','CHANGELOG.md','README.md','使用说明.md','使用说明.txt','版权与作者声明.txt')) {
    $p = Join-Path $src $f
    if (Test-Path -LiteralPath $p) { Copy-Item -LiteralPath $p -Destination $InstallDir -Force }
}
if (-not $NoDesktopShortcut) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $lnk = Join-Path $desktop 'CE转换引擎（Conversion Engine）.lnk'
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($lnk)
    $sc.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $sc.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $InstallDir 'app\pdfconv-launcher.ps1') + '"'
    $sc.WorkingDirectory = $InstallDir
    $sc.IconLocation = (Join-Path $InstallDir 'app\ce.ico') + ',0'
    $sc.Description = "CE转换引擎 v$version — 作者：温酒会知音"
    $sc.Save()
    Write-Host "已创建桌面快捷方式：$lnk"
}
Write-Host ""
Write-Host "安装完成。请双击桌面快捷方式启动。"
Write-Host "卸载：关闭软件后删除 $InstallDir 文件夹和桌面快捷方式。"
