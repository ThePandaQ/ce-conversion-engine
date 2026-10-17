<#
.SYNOPSIS
    在桌面创建 pdfconv 图形启动器快捷方式（可双击；支持把文件拖到图标上）。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\create-shortcut.ps1
#>
param(
    [string]$InstallDir = "$env:LOCALAPPDATA\PDFConv",
    [string]$Name = "CE转换引擎（Conversion Engine）"
)
$ErrorActionPreference = 'Stop'

$launcher = Join-Path $InstallDir 'app\pdfconv-launcher.ps1'
if (-not (Test-Path $launcher)) { throw "找不到启动器：$launcher（请先运行 install.ps1）" }

$desktop = [Environment]::GetFolderPath('Desktop')
$oldLnk = Join-Path $desktop 'PDF 转换 (pdfconv).lnk'
if (Test-Path $oldLnk) { Remove-Item $oldLnk -Force -ErrorAction SilentlyContinue }

$lnk = Join-Path $desktop ("$Name.lnk")
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($lnk)
$sc.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$sc.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$launcher`""
$sc.WorkingDirectory = $InstallDir
$ico = Join-Path $InstallDir 'app\ce.ico'
if (-not (Test-Path $ico)) { $ico = Join-Path $InstallDir 'app\pdfconv.ico' }
if (Test-Path $ico) { $sc.IconLocation = "$ico,0" }
$sc.Description = "CE转换引擎（Conversion Engine）：本地离线 PDF/OFD 转换（开源引擎）"
$sc.Save()
Write-Host "桌面快捷方式已创建: $lnk"
