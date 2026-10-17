<#
.SYNOPSIS
    运行 pdfconv 离线回归测试（会毒化代理，任何联网尝试都会失败）。
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\tests\offline_test.ps1
#>
param()
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $PSCommandPath
$proj = Split-Path -Parent $here
$venvPy = "$env:LOCALAPPDATA\PDFConv\venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) { throw "未找到安装的 venv：$venvPy，请先运行 install.ps1" }
Write-Host "== pdfconv offline test ==" -ForegroundColor Cyan
& $venvPy "$here\run_offline_test.py"
exit $LASTEXITCODE
