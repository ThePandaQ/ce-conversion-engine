<#
.SYNOPSIS
    pdfconv 桌面启动器（Windows Forms 图形界面）。

.DESCRIPTION
    双击桌面快捷方式即可打开；支持把 PDF / 图片拖到窗口或快捷方式图标上。
    选择转换方式后点击“开始转换”，全部转换在本机离线完成（开源引擎）。

    也可被命令行调用（供快捷方式拖放文件）：
        powershell -File pdfconv-launcher.ps1 "C:\a.pdf" "C:\b.pdf"
        powershell -File pdfconv-launcher.ps1 -SelfTest "C:\a.pdf"
#>
# 手工解析参数：避免 PowerShell 位置参数把“拖入的文件路径”误绑到 -Run 等命名参数。
$Files = New-Object System.Collections.Generic.List[string]
$SelfTest = $false
$SelfTestAsync = $false
$SelfTestSound = $false
$SelfTestLog = $false
$ScreenshotPath = ''
$ScreenshotSize = ''
$DemoLog = $false
$Run = ''
$RunOut = ''
$RunPages = ''
$RunLogFileArg = ''
$rawArgs = @($args)
for ($argi = 0; $argi -lt $rawArgs.Count; $argi++) {
    $a = [string]$rawArgs[$argi]
    if ($a -eq '-SelfTest') {
        $SelfTest = $true
    } elseif ($a -eq '-SelfTestAsync') {
        $SelfTestAsync = $true
    } elseif ($a -eq '-SelfTestSound') {
        $SelfTestSound = $true
    } elseif ($a -eq '-SelfTestLog') {
        $SelfTestLog = $true
    } elseif ($a -eq '-DemoLog') {
        $DemoLog = $true
    } elseif (($a -eq '-ScreenshotPath') -and ($argi + 1 -lt $rawArgs.Count)) {
        $ScreenshotPath = [string]$rawArgs[$argi + 1]; $argi++
    } elseif (($a -eq '-ScreenshotSize') -and ($argi + 1 -lt $rawArgs.Count)) {
        $ScreenshotSize = [string]$rawArgs[$argi + 1]; $argi++
    } elseif ((($a -eq '-Run') -or ($a -eq '-TestAction')) -and ($argi + 1 -lt $rawArgs.Count)) {
        $Run = [string]$rawArgs[$argi + 1]; $argi++
    } elseif (($a -eq '-RunLogFile') -and ($argi + 1 -lt $rawArgs.Count)) {
        $RunLogFileArg = [string]$rawArgs[$argi + 1]; $argi++
    } elseif (($a -eq '-RunOut') -and ($argi + 1 -lt $rawArgs.Count)) {
        $RunOut = [string]$rawArgs[$argi + 1]; $argi++
    } elseif (($a -eq '-RunPages') -and ($argi + 1 -lt $rawArgs.Count)) {
        $RunPages = [string]$rawArgs[$argi + 1]; $argi++
    } elseif ($a -like '-?*') {
        Write-Warning "忽略未知参数：$a"
    } elseif ($a) {
        [void]$Files.Add($a)
    }
}
$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    Write-Error "无法加载 Windows Forms：$($_.Exception.Message)"
    exit 1
}
[System.Windows.Forms.Application]::EnableVisualStyles()

$AppDir = $PSScriptRoot
$RootDir = Split-Path -Parent $AppDir
$Py = Join-Path $RootDir 'venv\Scripts\python.exe'
if (-not (Test-Path $Py)) {
    $portablePy = Join-Path $RootDir 'python\python.exe'
    if (Test-Path $portablePy) { $Py = $portablePy }
}
$Cli = Join-Path $AppDir 'pdfconv.py'
$HelpFile = Join-Path $AppDir 'README.md'
$IconFile = Join-Path $AppDir 'ce.ico'
if (-not (Test-Path $IconFile)) { $IconFile = Join-Path $AppDir 'pdfconv.ico' }
$BgFile = Join-Path $AppDir 'background.jpg'
$verFile = Join-Path $AppDir 'VERSION'
$script:AppVersion = '1.0.0'
if (Test-Path -LiteralPath $verFile) {
    try { $script:AppVersion = (Get-Content -LiteralPath $verFile -Raw -Encoding UTF8).Trim() } catch {}
}
$script:SettingsPath = Join-Path $RootDir 'settings.json'
$script:bgImage = $null
$script:coverImage = $null
$script:Settings = @{ sound = $true; background = $true; defaultOutDir = '' }

function Load-Settings {
    try {
        if (Test-Path -LiteralPath $script:SettingsPath) {
            $cfg = Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $props = @($cfg.PSObject.Properties.Name)
            if ($props -contains 'sound') { $script:Settings.sound = [bool]$cfg.sound }
            if ($props -contains 'background') { $script:Settings.background = [bool]$cfg.background }
            if ($props -contains 'defaultOutDir') { $script:Settings.defaultOutDir = [string]$cfg.defaultOutDir }
        }
    } catch {}
}

function Save-Settings {
    try {
        $json = $script:Settings | ConvertTo-Json -Compress
        [System.IO.File]::WriteAllText($script:SettingsPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}

$script:uiReady = $false
$script:ThemeBase = [System.Drawing.Color]::FromArgb(12, 32, 46)

function Mix-ThemeColor($photo, [double]$weight) {
    if (-not $photo) { return $script:ThemeBase }
    $b = $script:ThemeBase
    # 取照片区域色后整体压暗，贴近照片暗部并保证白字可读
    $r = [int][Math]::Min(255, (($photo.R * $weight + $b.R * (1 - $weight))) * 0.78)
    $g = [int][Math]::Min(255, (($photo.G * $weight + $b.G * (1 - $weight))) * 0.78)
    $bl = [int][Math]::Min(255, (($photo.B * $weight + $b.B * (1 - $weight))) * 0.78)
    return [System.Drawing.Color]::FromArgb($r, $g, $bl)
}

function Update-ThemeColors($bmp) {
    # 统一取整张背景的平均色，所有控件共用同一色，避免出现深浅不一的块状色差
    $photo = $null
    if ($bmp) {
        try {
            $thumb = New-Object System.Drawing.Bitmap($bmp, 32, 32)
            $sr = 0; $sg = 0; $sb = 0; $n = 0
            for ($yy = 0; $yy -lt $thumb.Height; $yy++) {
                for ($xx = 0; $xx -lt $thumb.Width; $xx++) {
                    $c = $thumb.GetPixel($xx, $yy)
                    $sr += $c.R; $sg += $c.G; $sb += $c.B; $n++
                }
            }
            $thumb.Dispose()
            if ($n -gt 0) {
                $photo = [System.Drawing.Color]::FromArgb([int]($sr / $n), [int]($sg / $n), [int]($sb / $n))
            }
        } catch {
            try { if ($thumb) { $thumb.Dispose() } } catch {}
        }
    }
    $baseColor = Mix-ThemeColor $photo 0.9
    try {
        if ($lstFiles) { $lstFiles.BackColor = $baseColor; $lstFiles.BorderStyle = 'None' }
        if ($cmbAction) { $cmbAction.BackColor = $baseColor; $cmbAction.FlatStyle = 'Flat' }
        if ($txtPages) { $txtPages.BackColor = $baseColor; $txtPages.BorderStyle = 'None' }
        if ($txtOut) { $txtOut.BackColor = $baseColor; $txtOut.BorderStyle = 'None' }
        foreach ($b in @($btnAddPdf, $btnAddImg, $btnClear, $btnBrowse, $btnSetDefault, $btnOpen, $btnHelp, $btnCancel)) {
            if ($b) {
                $b.BackColor = $baseColor
                $b.FlatAppearance.BorderSize = 0
            }
        }
        if ($btnRun) {
            # 主按钮：同色系里略亮一点，保持“开始转换”的强调感
            $accent = [System.Drawing.Color]::FromArgb(0, 112, 160)
            $r = [int](($baseColor.R * 0.55) + ($accent.R * 0.45))
            $g = [int](($baseColor.G * 0.55) + ($accent.G * 0.45))
            $bl = [int](($baseColor.B * 0.55) + ($accent.B * 0.45))
            $btnRun.BackColor = [System.Drawing.Color]::FromArgb($r, $g, $bl)
            $btnRun.FlatAppearance.BorderSize = 0
        }
    } catch {}
}

function Update-CoverBackground {
    # 按当前窗口尺寸裁切照片（cover 模式），保证照片始终铺满、不留黑边
    if ($Run) { return }
    if (-not $script:bgImage) { return }
    try {
        if ($chkBackground -and -not $chkBackground.Checked) {
            $form.BackgroundImage = $null
            return
        }        $cw = [Math]::Max(1, $form.ClientSize.Width)
        $ch = [Math]::Max(1, $form.ClientSize.Height)
        $iw = $script:bgImage.Width
        $ih = $script:bgImage.Height
        $scale = [Math]::Max($cw / [double]$iw, $ch / [double]$ih)
        $sw = [int][Math]::Ceiling($cw / $scale)
        $sh = [int][Math]::Ceiling($ch / $scale)
        if ($sw -gt $iw) { $sw = $iw }
        if ($sh -gt $ih) { $sh = $ih }
        $sx = [int](($iw - $sw) / 2)
        $sy = [int](($ih - $sh) / 2)
        $cover = New-Object System.Drawing.Bitmap($cw, $ch)
        $g = [System.Drawing.Graphics]::FromImage($cover)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $destRect = New-Object System.Drawing.Rectangle(0, 0, $cw, $ch)
        $srcRect = New-Object System.Drawing.Rectangle($sx, $sy, $sw, $sh)
        $g.DrawImage($script:bgImage, $destRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
        $g.Dispose()
        $old = $script:coverImage
        $script:coverImage = $cover
        $form.BackgroundImage = $cover
        $form.BackgroundImageLayout = 'Stretch'
        if ($old) { try { $old.Dispose() } catch {} }
        if ($script:uiReady) { Update-ThemeColors $cover }
    } catch {}
}

function Apply-Background {
    if ($Run) { return }   # 无界面子进程（转换任务）不需要背景图，省内存
    if (-not (Test-Path -LiteralPath $BgFile)) { return }
    try {
        if ($chkBackground -and -not $chkBackground.Checked) {
            $form.BackgroundImage = $null
            if ($script:uiReady) { Update-ThemeColors $null }
            return
        }
        if (-not $script:bgImage) {
            # 照片铺满整窗，叠加一层半透明黑色保证白色文字可读
            $src = [System.Drawing.Image]::FromFile($BgFile)
            $bmp = New-Object System.Drawing.Bitmap($src.Width, $src.Height)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.DrawImage($src, 0, 0, $src.Width, $src.Height)
            $veil = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(110, 0, 0, 0))
            $g.FillRectangle($veil, 0, 0, $bmp.Width, $bmp.Height)
            $g.Dispose(); $veil.Dispose(); $src.Dispose()
            $script:bgImage = $bmp
        }
        Update-CoverBackground
    } catch {}
}

Load-Settings

# ---------------------------------------------------------------- dark photo theme helpers
function Set-TransparentWhite($ctl) {
    try {
        $ctl.BackColor = [System.Drawing.Color]::Transparent
        $ctl.ForeColor = [System.Drawing.Color]::White
    } catch {}
}

function Set-DarkInput($ctl) {
    try {
        $ctl.BackColor = [System.Drawing.Color]::FromArgb(14, 36, 50)
        $ctl.ForeColor = [System.Drawing.Color]::White
        if ($ctl -is [System.Windows.Forms.TextBoxBase]) { $ctl.BorderStyle = 'FixedSingle' }
        if ($ctl -is [System.Windows.Forms.ComboBox]) { $ctl.FlatStyle = 'Flat' }
    } catch {}
}

function Set-DarkButton($btn, [bool]$primary = $false) {
    try {
        $btn.FlatStyle = 'Flat'
        if ($primary) {
            $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 112, 160)
        } else {
            $btn.BackColor = [System.Drawing.Color]::FromArgb(18, 46, 64)
        }
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.FlatAppearance.BorderSize = 0
        $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 150, 180)
    } catch {}
}

function Set-DarkTitleBar {
    # Windows 10/11：把系统标题栏改为深色，并与背景深色调一致
    if ($Run) { return }
    try {
        if (-not ('CEWin32.Dwm' -as [type])) {
            Add-Type -Namespace CEWin32 -Name Dwm -MemberDefinition '[DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);'
        }
        $hwnd = $form.Handle
        $dark = 1
        [void][CEWin32.Dwm]::DwmSetWindowAttribute($hwnd, 20, [ref]$dark, 4)
        $c = [System.Drawing.Color]::FromArgb(10, 30, 42)
        $caption = $c.R -bor ($c.G -shl 8) -bor ($c.B -shl 16)
        [void][CEWin32.Dwm]::DwmSetWindowAttribute($hwnd, 35, [ref]$caption, 4)
        $white = 16777215
        [void][CEWin32.Dwm]::DwmSetWindowAttribute($hwnd, 36, [ref]$white, 4)
    } catch {}
}

if (-not (Test-Path $Py)) {
    [System.Windows.Forms.MessageBox]::Show("找不到 Python 环境：`n$Py`n`n请重新运行 install.ps1 安装。", 'CE转换引擎', 'OK', 'Error') | Out-Null
    exit 1
}
if (-not (Test-Path $Cli)) {
    [System.Windows.Forms.MessageBox]::Show("找不到程序文件：`n$Cli`n`n请重新运行 install.ps1 安装。", 'CE转换引擎', 'OK', 'Error') | Out-Null
    exit 1
}

$LauncherPath = $PSCommandPath
$script:runProc = $null
$script:runStart = $null
$script:runOutFile = $null
$script:runErrFile = $null
$script:runLogFile = $null
$script:tickCount = 0
$script:cancelled = $false
$script:runTimer = New-Object System.Windows.Forms.Timer
$script:runTimer.Interval = 500

function Stop-RunProc {
    if ($script:runProc -and -not $script:runProc.HasExited) {
        try {
            & "$env:SystemRoot\System32\taskkill.exe" /PID $script:runProc.Id /T /F | Out-Null
        } catch {
            try { $script:runProc.Kill() } catch {}
        }
    }
}

# ---------------------------------------------------------------- helpers
function Quote-Arg([string]$a) {
    if ($null -eq $a) { return '""' }
    # 始终加引号（对 python / powershell 参数解析都安全），避免路径含空格、&、括号、中文标点等被拆分。
    return '"' + ($a -replace '"', '\"') + '"'
}

function Invoke-Pdfconv {
    param([string[]]$Argv)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Py
    # 关键：必须把 pdfconv.py 作为第一个参数传给 Python，否则 Python 会把子命令当脚本文件。
    $fullArgv = @($Cli) + $Argv
    $psi.Arguments = (($fullArgv | ForEach-Object { Quote-Arg $_ }) -join ' ')
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $psi.WorkingDirectory = $RootDir
    $p = [System.Diagnostics.Process]::Start($psi)
    $out = $p.StandardOutput.ReadToEnd()
    $err = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    return @{ ExitCode = $p.ExitCode; Out = $out; Err = $err; CmdLine = ($Py + ' ' + $psi.Arguments) }
}

function Get-UniquePath([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [System.IO.Path]::GetExtension($Path)
    for ($i = 1; $i -lt 1000; $i++) {
        $cand = Join-Path $dir ("$name-$i$ext")
        if (-not (Test-Path -LiteralPath $cand)) { return $cand }
    }
    return $Path
}

function Get-UniqueDir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }
    for ($i = 1; $i -lt 1000; $i++) {
        $cand = "$Path-$i"
        if (-not (Test-Path -LiteralPath $cand)) { return $cand }
    }
    return $Path
}

function Start-ConversionAsync {
    param(
        [hashtable]$Action,
        [string[]]$FileList,
        [string]$OutDir,
        [string]$Pages
    )
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logDir = Join-Path $RootDir 'logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Force -Path $logDir | Out-Null }
    $script:runOutFile = Join-Path $env:TEMP "pdfconv-$stamp.out.txt"
    $script:runErrFile = Join-Path $env:TEMP "pdfconv-$stamp.err.txt"
    $script:runLogFile = Join-Path $logDir "run-$stamp.log"
    $childArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
                   '-File', $LauncherPath, '-Run', [string]$Action.Cmd, '-RunOut', $OutDir,
                   '-RunLogFile', $script:runLogFile)
    if ($Pages) { $childArgs += @('-RunPages', $Pages) }
    $childArgs += $FileList
    $argLine = ($childArgs | ForEach-Object { Quote-Arg $_ }) -join ' '
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:runStart = Get-Date
    $script:tickCount = 0
    $script:cancelled = $false
    $script:runProc = Start-Process -FilePath $psExe -ArgumentList $argLine -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $script:runOutFile -RedirectStandardError $script:runErrFile
    $script:runTimer.Start()
}

# ---------------------------------------------------------------- actions
$actions = @(
    @{ Label = 'PDF → Word (.docx)';        Cmd = 'pdf2docx';       Out = 'file'; Ext = '.docx';            Pages = $true }
    @{ Label = 'PDF → Excel (.xlsx)';       Cmd = 'pdf2xlsx';       Out = 'file'; Ext = '.xlsx';            Pages = $true }
    @{ Label = 'PDF → PPT (.pptx)';         Cmd = 'pdf2pptx';       Out = 'file'; Ext = '.pptx';            Pages = $true }
    @{ Label = 'PDF → 图片 PNG';            Cmd = 'pdf2png';        Out = 'dir';  Suffix = '-png';          Pages = $true }
    @{ Label = 'PDF → 图片 JPG';            Cmd = 'pdf2jpg';        Out = 'dir';  Suffix = '-jpg';          Pages = $true }
    @{ Label = 'PDF → 文本 TXT';            Cmd = 'pdf2txt';        Out = 'file'; Ext = '.txt';             Pages = $true }
    @{ Label = 'PDF → 纯图/扫描型 PDF';     Cmd = 'pdf2imgpdf';     Out = 'file'; Ext = '-imgpdf.pdf';      Pages = $true }
    @{ Label = 'PDF → 可搜索 PDF (OCR)';    Cmd = 'pdf2searchable'; Out = 'file'; Ext = '-searchable.pdf';  Pages = $true }
    @{ Label = 'PDF → 提取内嵌图片';        Cmd = 'extract-images'; Out = 'dir';  Suffix = '-images';       Pages = $true }
    @{ Label = 'PDF → 提取页面';            Cmd = 'extract-pages';  Out = 'file'; Ext = '-pages.pdf';       Pages = $true; RequirePages = $true }
    @{ Label = 'PDF 拆分';                  Cmd = 'split';          Out = 'dir';  Suffix = '-split';        Pages = $false }
    @{ Label = 'PDF 合并';                  Cmd = 'merge';          Out = 'file'; Ext = '.pdf' }
    @{ Label = '图片 → PDF';                Cmd = 'img2pdf';        Out = 'file'; Ext = '.pdf' }
    @{ Label = '本地 OCR → 文本';           Cmd = 'ocr';            Out = 'file'; Ext = '.ocr.txt';         Pages = $true }
    @{ Label = 'PDF → OFD (.ofd)';          Cmd = 'pdf2ofd';        Out = 'file'; Ext = '.ofd' }
    @{ Label = 'OFD → PDF';                 Cmd = 'ofd2pdf';        Out = 'file'; Ext = '.pdf' }
    @{ Label = 'OFD → Word (.docx)';        Cmd = 'ofd2docx';       Out = 'file'; Ext = '.docx';            Pages = $true }
    @{ Label = 'OFD → Excel (.xlsx)';       Cmd = 'ofd2xlsx';       Out = 'file'; Ext = '.xlsx';            Pages = $true }
    @{ Label = 'OFD → PPT (.pptx)';         Cmd = 'ofd2pptx';       Out = 'file'; Ext = '.pptx';            Pages = $true }
    @{ Label = 'OFD → 图片 PNG';            Cmd = 'ofd2png';        Out = 'dir';  Suffix = '-png';          Pages = $true }
    @{ Label = 'OFD → 文本 TXT';            Cmd = 'ofd2txt';        Out = 'file'; Ext = '.txt';             Pages = $true }
    @{ Label = 'WPS → Word (.docx)';        Cmd = 'wps2docx';       Out = 'file'; Ext = '.docx' }
    @{ Label = 'WPS → PDF';                 Cmd = 'wps2pdf';        Out = 'file'; Ext = '.pdf' }
    @{ Label = 'WPS → 文本 TXT';            Cmd = 'wps2txt';        Out = 'file'; Ext = '.txt' }
    @{ Label = 'WPS → OFD (.ofd)';          Cmd = 'wps2ofd';        Out = 'file'; Ext = '.ofd' }
    @{ Label = 'Word/RTF/TXT → WPS (.wps)'; Cmd = 'doc2wps';        Out = 'file'; Ext = '.wps' }
    @{ Label = 'PDF → WPS (.wps)';          Cmd = 'pdf2wps';        Out = 'file'; Ext = '.wps' }
    @{ Label = 'OFD → WPS (.wps)';          Cmd = 'ofd2wps';        Out = 'file'; Ext = '.wps' }
)

# ---------------------------------------------------------------- conversion engine (UI 与无界面模式共用)
function Invoke-Conversion {
    param(
        [hashtable]$Action,
        [string[]]$FileList,
        [string]$OutDir,
        [string]$Pages,
        [System.Text.StringBuilder]$Log
    )
    $ok = 0
    $fail = 0
    if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

    $runOne = {
        param([string[]]$Argv)
        $r = Invoke-Pdfconv $Argv
        [void]$Log.AppendLine('> ' + $r.CmdLine)
        if ($r.Out) { [void]$Log.AppendLine($r.Out.TrimEnd()) }
        if ($r.Err) { [void]$Log.AppendLine($r.Err.TrimEnd()) }
        return $r.ExitCode
    }

    if ($Action.Cmd -eq 'merge') {
        $out = Get-UniquePath (Join-Path $OutDir 'merged.pdf')
        $argv = @('merge') + $FileList + @('-o', $out, '--overwrite')
        if ((& $runOne $argv) -eq 0) { $ok++ } else { $fail++ }
    }
    elseif ($Action.Cmd -eq 'img2pdf') {
        $out = Get-UniquePath (Join-Path $OutDir 'images.pdf')
        $argv = @('img2pdf') + $FileList + @('-o', $out, '--overwrite')
        if ((& $runOne $argv) -eq 0) { $ok++ } else { $fail++ }
    }
    else {
        foreach ($f in $FileList) {
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($f)
            $argv = @($Action.Cmd, $f)
            if ($Action.Out -eq 'file') {
                $out = Get-UniquePath (Join-Path $OutDir ($stem + $Action.Ext))
            } else {
                $out = Get-UniqueDir (Join-Path $OutDir ($stem + $Action.Suffix))
            }
            $argv += @('-o', $out)
            if ($Action.Cmd -eq 'split') {
                if ($Pages) { $argv += @('--ranges', $Pages) } else { $argv += @('--every', '1') }
            } elseif ($Pages -and $Action.Pages) {
                $argv += @('--pages', $Pages)
            } elseif ($Action.RequirePages -and -not $Pages) {
                [void]$Log.AppendLine("跳过（该操作需要填写页码）：$f")
                $fail++
                continue
            }
            $argv += '--overwrite'
            if ((& $runOne $argv) -eq 0) { $ok++ } else { $fail++ }
        }
    }
    return @{ Ok = $ok; Fail = $fail; OutDir = $OutDir }
}

# ---------------------------------------------------------------- form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CE转换引擎（Conversion Engine） v$($script:AppVersion) — 作者：温酒会知音"
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(860, 680)
$form.MinimumSize = New-Object System.Drawing.Size(820, 620)
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(10, 30, 42)
$form.ForeColor = [System.Drawing.Color]::White
if (Test-Path $IconFile) {
    try { $form.Icon = New-Object System.Drawing.Icon($IconFile) } catch {}
}

# 标题栏改用深色（Windows 10/11 DWM），与照片背景协调
$form.Add_HandleCreated({ Set-DarkTitleBar })

# 窗口缩放时重新裁切照片，保证整窗始终被照片覆盖
$script:bgResizeTimer = New-Object System.Windows.Forms.Timer
$script:bgResizeTimer.Interval = 120
$script:bgResizeTimer.Add_Tick({
    $script:bgResizeTimer.Stop()
    Update-CoverBackground
})
$form.Add_Resize({
    if (-not $Run) {
        $script:bgResizeTimer.Stop()
        $script:bgResizeTimer.Start()
    }
})

$grpInput = New-Object System.Windows.Forms.Panel
$grpInput.Location = New-Object System.Drawing.Point(12, 8)
$grpInput.Size = New-Object System.Drawing.Size(836, 148)
$grpInput.Anchor = 'Top,Left,Right'
$grpInput.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($grpInput)

$lblInputHeader = New-Object System.Windows.Forms.Label
$lblInputHeader.Text = '输入文件（可拖入 PDF / 图片）'
$lblInputHeader.Location = New-Object System.Drawing.Point(12, 4)
$lblInputHeader.AutoSize = $true
$lblInputHeader.ForeColor = [System.Drawing.Color]::White
$lblInputHeader.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9, [System.Drawing.FontStyle]::Bold)
$grpInput.Controls.Add($lblInputHeader)

$lstFiles = New-Object System.Windows.Forms.ListBox
$lstFiles.Location = New-Object System.Drawing.Point(12, 26)
$lstFiles.Size = New-Object System.Drawing.Size(636, 110)
$lstFiles.SelectionMode = 'MultiExtended'
$lstFiles.AllowDrop = $true
$lstFiles.Anchor = 'Top,Left,Right'
$grpInput.Controls.Add($lstFiles)

$btnAddPdf = New-Object System.Windows.Forms.Button
$btnAddPdf.Text = '添加文档...'
$btnAddPdf.Location = New-Object System.Drawing.Point(660, 26)
$btnAddPdf.Size = New-Object System.Drawing.Size(160, 30)
$btnAddPdf.Anchor = 'Top,Right'
$grpInput.Controls.Add($btnAddPdf)

$btnAddImg = New-Object System.Windows.Forms.Button
$btnAddImg.Text = '添加图片...'
$btnAddImg.Location = New-Object System.Drawing.Point(660, 62)
$btnAddImg.Size = New-Object System.Drawing.Size(160, 30)
$btnAddImg.Anchor = 'Top,Right'
$grpInput.Controls.Add($btnAddImg)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = '清空列表'
$btnClear.Location = New-Object System.Drawing.Point(660, 98)
$btnClear.Size = New-Object System.Drawing.Size(160, 30)
$btnClear.Anchor = 'Top,Right'
$grpInput.Controls.Add($btnClear)

$lblAction = New-Object System.Windows.Forms.Label
$lblAction.Text = '转换方式：'
$lblAction.Location = New-Object System.Drawing.Point(12, 170)
$lblAction.AutoSize = $true
$form.Controls.Add($lblAction)

$cmbAction = New-Object System.Windows.Forms.ComboBox
$cmbAction.Location = New-Object System.Drawing.Point(92, 166)
$cmbAction.Size = New-Object System.Drawing.Size(400, 26)
$cmbAction.DropDownStyle = 'DropDownList'
foreach ($a in $actions) { [void]$cmbAction.Items.Add($a.Label) }
$cmbAction.SelectedIndex = 0
$form.Controls.Add($cmbAction)

$lblPages = New-Object System.Windows.Forms.Label
$lblPages.Text = '页码/范围：'
$lblPages.Location = New-Object System.Drawing.Point(508, 170)
$lblPages.AutoSize = $true
$form.Controls.Add($lblPages)

$txtPages = New-Object System.Windows.Forms.TextBox
$txtPages.Location = New-Object System.Drawing.Point(592, 166)
$txtPages.Size = New-Object System.Drawing.Size(110, 26)
$txtPages.Text = ''
$form.Controls.Add($txtPages)

$lblPagesHint = New-Object System.Windows.Forms.Label
$lblPagesHint.Text = '留空=全部'
$lblPagesHint.Location = New-Object System.Drawing.Point(710, 170)
$lblPagesHint.AutoSize = $true
$lblPagesHint.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblPagesHint)

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = '输出目录：'
$lblOut.Location = New-Object System.Drawing.Point(12, 206)
$lblOut.AutoSize = $true
$form.Controls.Add($lblOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(92, 202)
$txtOut.Size = New-Object System.Drawing.Size(520, 26)
$txtOut.Text = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'pdfconv-output')
$txtOut.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtOut)
if ($script:Settings.defaultOutDir) { $txtOut.Text = [string]$script:Settings.defaultOutDir }

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '浏览...'
$btnBrowse.Location = New-Object System.Drawing.Point(736, 201)
$btnBrowse.Size = New-Object System.Drawing.Size(112, 28)
$btnBrowse.Anchor = 'Top,Right'
$form.Controls.Add($btnBrowse)

$btnSetDefault = New-Object System.Windows.Forms.Button
$btnSetDefault.Text = '设为默认'
$btnSetDefault.Location = New-Object System.Drawing.Point(620, 201)
$btnSetDefault.Size = New-Object System.Drawing.Size(108, 28)
$btnSetDefault.Anchor = 'Top,Right'
$form.Controls.Add($btnSetDefault)
$btnSetDefault.Add_Click({
    $dir = $txtOut.Text.Trim()
    if (-not $dir) { Show-Error '输出目录为空，无法设为默认。'; return }
    $script:Settings.defaultOutDir = $dir
    Save-Settings
    $status.Text = "已设为默认输出目录：$dir"
    $status.ForeColor = [System.Drawing.Color]::FromArgb(220, 232, 240)
})

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '开始转换'
$btnRun.Location = New-Object System.Drawing.Point(12, 240)
$btnRun.Size = New-Object System.Drawing.Size(160, 38)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(31, 78, 121)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = 'Flat'
$form.Controls.Add($btnRun)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = '打开输出目录'
$btnOpen.Location = New-Object System.Drawing.Point(180, 240)
$btnOpen.Size = New-Object System.Drawing.Size(140, 38)
$form.Controls.Add($btnOpen)

$btnHelp = New-Object System.Windows.Forms.Button
$btnHelp.Text = '命令帮助'
$btnHelp.Location = New-Object System.Drawing.Point(328, 240)
$btnHelp.Size = New-Object System.Drawing.Size(110, 38)
$form.Controls.Add($btnHelp)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = '取消'
$btnCancel.Location = New-Object System.Drawing.Point(446, 240)
$btnCancel.Size = New-Object System.Drawing.Size(96, 38)
$form.Controls.Add($btnCancel)

$chkSound = New-Object System.Windows.Forms.CheckBox
$chkSound.Text = '提示音'
$chkSound.Location = New-Object System.Drawing.Point(548, 236)
$chkSound.Size = New-Object System.Drawing.Size(74, 20)
$chkSound.Checked = [bool]$script:Settings.sound
$form.Controls.Add($chkSound)
$chkSound.Add_CheckedChanged({
    $script:Settings.sound = [bool]$chkSound.Checked
    Save-Settings
})

$chkBackground = New-Object System.Windows.Forms.CheckBox
$chkBackground.Text = '背景'
$chkBackground.Location = New-Object System.Drawing.Point(548, 258)
$chkBackground.Size = New-Object System.Drawing.Size(74, 20)
$chkBackground.Checked = [bool]$script:Settings.background
$form.Controls.Add($chkBackground)
$chkBackground.Add_CheckedChanged({
    $script:Settings.background = [bool]$chkBackground.Checked
    Save-Settings
    Apply-Background
})

Apply-Background

$status = New-Object System.Windows.Forms.Label
$status.Text = '就绪。作者：温酒会知音 · 全部转换在本机离线完成。'
$status.Location = New-Object System.Drawing.Point(12, 280)
$status.Size = New-Object System.Drawing.Size(836, 32)
$status.AutoSize = $false
$status.TextAlign = 'MiddleLeft'
$status.ForeColor = [System.Drawing.Color]::FromArgb(220, 232, 240)
$status.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($status)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = '输出 / 日志（绿色=成功，红色=失败；双击可打开日志文件）：'
$lblResult.Location = New-Object System.Drawing.Point(12, 320)
$lblResult.AutoSize = $true
$form.Controls.Add($lblResult)

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Location = New-Object System.Drawing.Point(12, 342)
$pnlLog.Size = New-Object System.Drawing.Size(836, 324)
$pnlLog.Anchor = 'Top,Bottom,Left,Right'
$pnlLog.AutoScroll = $true
$pnlLog.BackColor = [System.Drawing.Color]::Transparent
$pnlLog.BorderStyle = 'FixedSingle'
$form.Controls.Add($pnlLog)

$script:logLines = @()
$script:logFont = New-Object System.Drawing.Font('Consolas', 9)
$script:lastLogText = ''
$script:renderingLog = $false
$script:LogColors = @{
    Ok     = [System.Drawing.Color]::FromArgb(0, 230, 130)      # 绿：成功 / [OK] / [DONE] fail=0
    Err    = [System.Drawing.Color]::FromArgb(255, 90, 90)      # 红：失败 / [FAIL] / [ERROR] / 异常
    Warn   = [System.Drawing.Color]::FromArgb(255, 190, 70)     # 橙：警告
    Normal = [System.Drawing.Color]::FromArgb(228, 236, 242)    # 正文（浅色）
}

function Get-LogColorKind([string]$line) {
    if ($line -match '\[FAIL\]|\[ERROR\]|Traceback|Exception|错误|失败|\[DONE\].*fail=[1-9]') { return 'err' }
    elseif ($line -match '\[OK\]|\[DONE\].*fail=0|成功') { return 'ok' }
    elseif ($line -match '^\s*warning:|警告') { return 'warn' }
    return 'normal'
}

function Render-LogText([string]$text) {
    # 日志按面板宽度自动换行；只保留右侧竖向滚动条
    if ($script:renderingLog) { return }
    $script:renderingLog = $true
    $pnlLog.SuspendLayout()
    try {
        $pnlLog.Controls.Clear()
        $script:logLines = @()
        $wrapWidth = [Math]::Max(120, $pnlLog.ClientSize.Width - 26)
        $lines = ([string]$text) -split "`r?`n"
        $y = 4
        foreach ($line in $lines) {
            $script:logLines += $line
            $color = switch (Get-LogColorKind $line) {
                'ok' { $script:LogColors.Ok }
                'err' { $script:LogColors.Err }
                'warn' { $script:LogColors.Warn }
                default { $script:LogColors.Normal }
            }
            $chunks = @()
            if ($line.Length -eq 0) {
                $chunks = @('')
            } else {
                $buf = ''
                foreach ($ch in $line.ToCharArray()) {
                    $cand = $buf + $ch
                    if ($buf.Length -gt 0 -and [System.Windows.Forms.TextRenderer]::MeasureText($cand, $script:logFont).Width -gt $wrapWidth) {
                        $chunks += $buf
                        $buf = [string]$ch
                    } else {
                        $buf = $cand
                    }
                }
                if ($buf.Length -gt 0) { $chunks += $buf }
            }
            foreach ($chunk in $chunks) {
                $lbl = New-Object System.Windows.Forms.Label
                $lbl.AutoSize = $false
                $lbl.Left = 4
                $lbl.Top = $y
                $lbl.Height = 18
                $lbl.Width = $wrapWidth
                $lbl.Text = $chunk
                $lbl.Font = $script:logFont
                $lbl.ForeColor = $color
                $lbl.BackColor = [System.Drawing.Color]::Transparent
                $lbl.UseMnemonic = $false
                $pnlLog.Controls.Add($lbl)
                $y += 18
            }
        }
        $pnlLog.AutoScrollMinSize = New-Object System.Drawing.Size(0, ($y + 2))
    } finally {
        $pnlLog.ResumeLayout()
        $script:renderingLog = $false
    }
}

function Set-LogText([string]$text) {
    $script:lastLogText = [string]$text
    Render-LogText $script:lastLogText
}

# 面板宽度变化时重新换行（避免出现横向滚动条）
$pnlLog.Add_Resize({
    if (-not $script:renderingLog -and $script:lastLogText) {
        Render-LogText $script:lastLogText
    }
})

function Get-LogText {
    if ($script:logLines -and $script:logLines.Count -gt 0) { return ($script:logLines -join "`r`n") }
    return ''
}

$pnlLog.Add_DoubleClick({
    if ($script:runLogFile -and (Test-Path -LiteralPath $script:runLogFile)) {
        try { Start-Process notepad.exe -ArgumentList ('"' + $script:runLogFile + '"') } catch {}
    }
})

$menuLog = New-Object System.Windows.Forms.ContextMenuStrip
$miCopyLog = $menuLog.Items.Add('复制全部日志')
$miCopyLog.Add_Click({ try { [System.Windows.Forms.Clipboard]::SetText((Get-LogText)) } catch {} })
$miAbout = $menuLog.Items.Add('关于 / 版权')
$miAbout.Add_Click({
    $msg = "CE转换引擎（Conversion Engine） v$($script:AppVersion)`r`n作者：温酒会知音`r`n版权所有 © 2026 温酒会知音`r`n许可：AGPL-3.0（详见 LICENSE）`r`n`r`n本软件在本机离线运行，不上传任何文件。`r`n.wps 格式转换需本机安装 WPS Office。`r`n`r`n第三方依赖与许可：见 NOTICE.md"
    [System.Windows.Forms.MessageBox]::Show($msg, '关于 / 版权', 'OK', 'Information') | Out-Null
})
$miOpenLog = $menuLog.Items.Add('打开日志文件')
$miOpenLog.Add_Click({ if ($script:runLogFile -and (Test-Path -LiteralPath $script:runLogFile)) { try { Start-Process notepad.exe -ArgumentList ('"' + $script:runLogFile + '"') } catch {} } })
$pnlLog.ContextMenuStrip = $menuLog

# ---------------------------------------------------------------- dark photo theme
foreach ($lbl in @($lblAction, $lblPages, $lblOut, $lblResult, $lblInputHeader)) { Set-TransparentWhite $lbl }
$lblPagesHint.BackColor = [System.Drawing.Color]::Transparent
$lblPagesHint.ForeColor = [System.Drawing.Color]::FromArgb(190, 210, 220)
foreach ($chk in @($chkSound, $chkBackground)) {
    $chk.BackColor = [System.Drawing.Color]::Transparent
    $chk.ForeColor = [System.Drawing.Color]::White
}
foreach ($inp in @($lstFiles, $txtPages, $txtOut, $cmbAction)) { Set-DarkInput $inp }
foreach ($b in @($btnAddPdf, $btnAddImg, $btnClear, $btnBrowse, $btnSetDefault, $btnOpen, $btnHelp, $btnCancel)) { Set-DarkButton $b }
Set-DarkButton $btnRun $true

# 界面就绪后按背景照片为控件做自适应配色
$script:uiReady = $true
if ($script:coverImage) { Update-ThemeColors $script:coverImage } else { Update-ThemeColors $null }

# ---------------------------------------------------------------- state
$script:lastOutDir = $txtOut.Text

function Add-Files([string[]]$paths) {
    foreach ($p in $paths) {
        if ($p -and (Test-Path -LiteralPath $p)) {
            $full = (Resolve-Path -LiteralPath $p).ProviderPath
            if (-not $lstFiles.Items.Contains($full)) { [void]$lstFiles.Items.Add($full) }
        }
    }
}

function Show-Error([string]$msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, 'CE转换引擎', 'OK', 'Error') | Out-Null
}

function Play-ResultSound([bool]$success) {
    # 成功：两声音高上行；失败：两声音高下行（不依赖系统声音方案）
    try {
        if ($chkSound -and -not $chkSound.Checked) { return }
    } catch { return }
    try {
        if ($success) {
            [Console]::Beep(880, 110)
            Start-Sleep -Milliseconds 60
            [Console]::Beep(1318, 170)
        } else {
            [Console]::Beep(392, 200)
            Start-Sleep -Milliseconds 70
            [Console]::Beep(262, 320)
        }
    } catch {
        # 无控制台/蜂鸣受限时退回系统提示音
        try {
            if ($success) { [System.Media.SystemSounds]::Asterisk.Play() }
            else { [System.Media.SystemSounds]::Hand.Play() }
        } catch {}
    }
}

# ---------------------------------------------------------------- event handlers
$btnAddPdf.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '选择待转换文档（PDF/OFD/WPS/Word/RTF/TXT）'
    $dlg.Filter = '文档 (*.pdf;*.ofd;*.wps;*.doc;*.docx;*.rtf;*.txt)|*.pdf;*.ofd;*.wps;*.doc;*.docx;*.rtf;*.txt|PDF/OFD (*.pdf;*.ofd)|*.pdf;*.ofd|WPS/Word (*.wps;*.doc;*.docx;*.rtf;*.txt)|*.wps;*.doc;*.docx;*.rtf;*.txt|所有文件 (*.*)|*.*'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq 'OK') { Add-Files $dlg.FileNames }
})

$btnAddImg.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '选择图片文件'
    $dlg.Filter = '图片 (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp)|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp|所有文件 (*.*)|*.*'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq 'OK') { Add-Files $dlg.FileNames }
})

$btnClear.Add_Click({ $lstFiles.Items.Clear() })

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择输出目录'
    if (Test-Path -LiteralPath $txtOut.Text) { $dlg.SelectedPath = $txtOut.Text }
    if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.SelectedPath }
})

$btnOpen.Add_Click({
    if ($script:lastOutDir -and (Test-Path -LiteralPath $script:lastOutDir)) {
        Start-Process explorer.exe -ArgumentList ('"' + $script:lastOutDir + '"')
    } else {
        $status.Text = "还没有可打开的输出目录（完成一次转换后再试）"
        $status.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 235)
    }
})

$btnHelp.Add_Click({
    if (Test-Path $HelpFile) { Start-Process notepad.exe -ArgumentList ('"' + $HelpFile + '"') }
    else {
        $r = Invoke-Pdfconv @('--help')
        Set-LogText ($r.Out + "`r`n" + $r.Err)
    }
})

$lstFiles.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } else { $e.Effect = 'None' }
})
$lstFiles.Add_DragDrop({
    param($sender, $e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Add-Files $paths
})
$form.AllowDrop = $true
$form.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } else { $e.Effect = 'None' }
})
$form.Add_DragDrop({
    param($sender, $e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Add-Files $paths
})

$script:runTimer.Add_Tick({
    try {
        $script:tickCount++
        if (-not $script:runProc) { $script:runTimer.Stop(); return }
        if (-not $script:runProc.HasExited) {
            $elapsed = [int]((Get-Date) - $script:runStart).TotalSeconds
            $status.Text = "转换中... 已运行 ${elapsed}s（大文件可能较久，窗口保持响应，可点取消按钮）"
            $status.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 235)
            return
        }
        $script:runTimer.Stop()
        $p = $script:runProc
        $script:runProc = $null
        $out = ''
        if ($script:runLogFile -and (Test-Path -LiteralPath $script:runLogFile)) {
            $out = Get-Content -LiteralPath $script:runLogFile -Raw -Encoding UTF8
        } elseif ($script:runOutFile -and (Test-Path -LiteralPath $script:runOutFile)) {
            $out = Get-Content -LiteralPath $script:runOutFile -Raw -Encoding UTF8
        }
        if ($out) {
            $esc = [string][char]27
            $out = $out -replace ([regex]::Escape($esc) + '\[[0-9;?]*[A-Za-z]'), ''
        }
        $err = ''
        if ($script:runErrFile -and (Test-Path -LiteralPath $script:runErrFile)) {
            $err = Get-Content -LiteralPath $script:runErrFile -Raw -Encoding UTF8
        }
        Set-LogText (($out + "`r`n" + $err).Trim())
        Remove-Item -LiteralPath $script:runOutFile, $script:runErrFile -Force -ErrorAction SilentlyContinue
        try {
            $logDir = Join-Path $RootDir 'logs'
            Get-ChildItem -LiteralPath $logDir -Filter 'run-*.log' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
                Remove-Item -Force -ErrorAction SilentlyContinue
        } catch {}

        if ($script:cancelled) {
            $status.Text = '已取消（可能已有部分输出）'
            $status.ForeColor = $script:LogColors.Warn
            $script:cancelled = $false
            $btnRun.Enabled = $true
            $btnRun.Text = '开始转换'
            return
        }

        $m = [regex]::Match([string]$out, '\[DONE\] ok=(\d+) fail=(\d+)')
        $success = $false
        if ($m.Success) {
            $okCount = [int]$m.Groups[1].Value
            $failCount = [int]$m.Groups[2].Value
            $status.Text = "完成：成功 $okCount，失败 $failCount。输出目录：$($txtOut.Text.Trim())"
            $success = ($okCount -gt 0 -and $failCount -eq 0)
            if ($success) { $status.ForeColor = $script:LogColors.Ok } else { $status.ForeColor = $script:LogColors.Err }
        } else {
            $status.Text = "转换进程已结束（退出码 $($p.ExitCode)）。详情见日志区。"
            $status.ForeColor = $script:LogColors.Err
        }
        $script:lastOutDir = $txtOut.Text.Trim()
        Play-ResultSound $success
        if (-not $success) {
            Show-Error ("转换失败，请查看日志区的完整命令与错误：`r`n`r`n" + (Get-LogText))
        }
        $btnRun.Enabled = $true
        $btnRun.Text = '开始转换'
    }
    catch {
        $script:runTimer.Stop()
        $script:runProc = $null
        $btnRun.Enabled = $true
        $btnRun.Text = '开始转换'
        Play-ResultSound $false
        $status.Text = '出错。'
        $status.ForeColor = $script:LogColors.Err
        Show-Error ("读取转换结果出错：`n" + $_.Exception.Message)
    }
})

$btnRun.Add_Click({
    try {
        $files = @($lstFiles.Items | ForEach-Object { [string]$_ })
        if ($files.Count -eq 0) { Show-Error '请先添加文件（可把 PDF/图片拖入列表）。'; return }
        if ($script:runProc -and -not $script:runProc.HasExited) { Show-Error '上一次转换还在进行中，请稍候。'; return }
        $action = $actions[$cmbAction.SelectedIndex]
        $outdir = $txtOut.Text.Trim()
        if (-not $outdir) { $outdir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'pdfconv-output' }
        $pages = $txtPages.Text.Trim()

        $btnRun.Enabled = $false
        $btnRun.Text = '转换中...'
        $script:lastOutDir = $outdir
        Set-LogText ('== CE转换引擎 桌面启动器 ==' + "`r`n" + '已开始转换，请稍候...（大文件可能需要数分钟，窗口保持响应）')
        Start-ConversionAsync -Action $action -FileList $files -OutDir $outdir -Pages $pages
        $status.Text = '转换中... 0s'
        $status.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 235)
    }
    catch {
        Show-Error ("启动转换失败：`n" + $_.Exception.Message)
        $btnRun.Enabled = $true
        $btnRun.Text = '开始转换'
        $status.Text = '出错了。'
        $status.ForeColor = $script:LogColors.Err
    }
})

$btnCancel.Add_Click({
    if ($script:runProc -and -not $script:runProc.HasExited) {
        $script:cancelled = $true
        Stop-RunProc
        $status.Text = '正在取消...'
        $status.ForeColor = $script:LogColors.Warn
    } else {
        $status.Text = '当前没有正在进行的转换。'
        $status.ForeColor = [System.Drawing.Color]::FromArgb(210, 225, 235)
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    Stop-RunProc
})

# ---------------------------------------------------------------- startup args / self test / headless run
foreach ($f in $Files) {
    if ($f -and (Test-Path -LiteralPath $f)) { Add-Files @($f) }
}

if ($SelfTestAsync) {
    $idx = 0
    if ($Run) {
        $idx = -1
        if ($Run -match '^\d+$') {
            $i = [int]$Run
            if ($i -ge 0 -and $i -lt $actions.Count) { $idx = $i }
        } else {
            for ($i = 0; $i -lt $actions.Count; $i++) {
                if ($actions[$i].Cmd -eq $Run -or $actions[$i].Label -eq $Run) { $idx = $i; break }
            }
        }
        if ($idx -lt 0) { [Console]::Error.WriteLine("未知动作：$Run"); exit 2 }
    }
    $headFiles = @($lstFiles.Items | ForEach-Object { [string]$_ })
    if ($headFiles.Count -eq 0) { [Console]::Error.WriteLine('没有输入文件。'); exit 2 }
    $headOut = $RunOut
    if (-not $headOut) { $headOut = Join-Path $env:TEMP 'pdfconv-selftest-async' }
    if (-not (Test-Path -LiteralPath $headOut)) { New-Item -ItemType Directory -Force -Path $headOut | Out-Null }
    Start-ConversionAsync -Action $actions[$idx] -FileList $headFiles -OutDir $headOut -Pages $RunPages
    $deadline = (Get-Date).AddSeconds(300)
    $lastTick = -1
    while ($script:runProc -and -not $script:runProc.HasExited -and (Get-Date) -lt $deadline) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        $elapsed = [int]((Get-Date) - $script:runStart).TotalSeconds
        if ($elapsed -ne $lastTick -and $elapsed % 5 -eq 0) {
            $lastTick = $elapsed
            [Console]::Error.WriteLine("...running ${elapsed}s")
        }
    }
    if (-not $script:runProc) { [Console]::Error.WriteLine('ASYNC FAIL: no child process'); exit 1 }
    if (-not $script:runProc.HasExited) {
        try { $script:runProc.Kill() } catch {}
        [Console]::Error.WriteLine('ASYNC FAIL: timeout'); exit 1
    }
    Start-Sleep -Milliseconds 300
    $childOut = ''
    if ($script:runLogFile -and (Test-Path -LiteralPath $script:runLogFile)) {
        $childOut = Get-Content -LiteralPath $script:runLogFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    } else {
        $childOut = Get-Content -LiteralPath $script:runOutFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    $childErr = Get-Content -LiteralPath $script:runErrFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    Set-LogText (($childOut + "`r`n" + $childErr).Trim())
    [Console]::Out.Write((Get-LogText))
    [Console]::Out.WriteLine('')
    $dm = [regex]::Match([string]$childOut, '\[DONE\] ok=(\d+) fail=(\d+)')
    $exitCode = 1
    if ($dm.Success -and [int]$dm.Groups[1].Value -gt 0 -and [int]$dm.Groups[2].Value -eq 0) { $exitCode = 0 }
    Remove-Item -LiteralPath $script:runOutFile, $script:runErrFile -Force -ErrorAction SilentlyContinue
    $ticks = $script:tickCount
    [Console]::Out.WriteLine("SELFTEST-ASYNC ticks=$ticks exit=$exitCode actions=$($actions.Count) outdir=$headOut")
    if ($exitCode -eq 0 -and $ticks -gt 0) { exit 0 } else { exit 1 }
}

if ($Run) {
    $idx = -1
    if ($Run -match '^\d+$') {
        $i = [int]$Run
        if ($i -ge 0 -and $i -lt $actions.Count) { $idx = $i }
    } else {
        for ($i = 0; $i -lt $actions.Count; $i++) {
            if ($actions[$i].Cmd -eq $Run -or $actions[$i].Label -eq $Run) { $idx = $i; break }
        }
    }
    if ($idx -lt 0) {
        [Console]::Error.WriteLine("未知动作：$Run（可用：0..$($actions.Count - 1) 或 pdf2docx/pdf2png/merge/...）")
        exit 2
    }
    $headFiles = @($lstFiles.Items | ForEach-Object { [string]$_ })
    if ($headFiles.Count -eq 0) {
        [Console]::Error.WriteLine('没有输入文件。用法：-Run pdf2docx [-RunOut DIR] [-RunPages 1-3] "file1" "file2" ...')
        exit 2
    }
    $headOut = $RunOut
    if (-not $headOut) { $headOut = Join-Path ([Environment]::GetFolderPath('Desktop')) 'pdfconv-output' }
    $headLog = New-Object System.Text.StringBuilder
    [void]$headLog.AppendLine('== CE转换引擎 转换日志 ==')
    [void]$headLog.AppendLine('时间: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    [void]$headLog.AppendLine('动作: ' + $actions[$idx].Label)
    [void]$headLog.AppendLine('输出: ' + $headOut)
    if ($RunLogFileArg) { [void]$headLog.AppendLine('日志: ' + $RunLogFileArg) }
    [void]$headLog.AppendLine('输入:')
    foreach ($hf in $headFiles) { [void]$headLog.AppendLine('  - ' + $hf) }
    [void]$headLog.AppendLine('')
    $headRes = Invoke-Conversion -Action $actions[$idx] -FileList $headFiles -OutDir $headOut -Pages $RunPages -Log $headLog
    [void]$headLog.AppendLine("[DONE] ok=$($headRes.Ok) fail=$($headRes.Fail) outdir=$($headRes.OutDir)")
    $headText = $headLog.ToString()
    $esc = [string][char]27
    $headText = $headText -replace ([regex]::Escape($esc) + '\[[0-9;?]*[A-Za-z]'), ''
    # 过滤第三方库的已知无害告警（PyMuPDF 旧别名 fitz 的弃用提示）
    $headText = $headText -replace '(?m)^warning: The `fitz` API is deprecated.*(\r?\n)?', ''
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    if ($RunLogFileArg) {
        try {
            $logParent = Split-Path -Parent $RunLogFileArg
            if ($logParent -and -not (Test-Path -LiteralPath $logParent)) { New-Item -ItemType Directory -Force -Path $logParent | Out-Null }
            [System.IO.File]::WriteAllText($RunLogFileArg, $headText, $utf8NoBom)
        } catch {}
    }
    [Console]::Out.Write($headText)
    if ($headRes.Ok -gt 0 -and $headRes.Fail -eq 0) { exit 0 } else { exit 1 }
}

if ($ScreenshotPath) {
    $form.CreateControl()
    [void]$form.Handle
    if ($DemoLog) {
        Set-LogText ((@(
            '== CE转换引擎 转换日志 ==',
            '时间: 2026-10-17 20:00:00',
            '动作: PDF → Word (.docx)',
            '> C:\Users\Public\AppData\Local\PDFConv\venv\Scripts\python.exe "C:\Users\Public\AppData\Local\PDFConv\app\pdfconv.py" "pdf2docx" "E:\文件\示例文档.docx" "-o" "E:\文件\文件格式转换\示例文档.docx" "--overwrite"',
            '[OK] pdf2docx: E:\文件.pdf -> E:\文件.docx (2.7s)',
            '[DONE] ok=1 fail=0 outdir=E:\文件格式转换',
            '[FAIL] pdf2xlsx: E:\坏文件.pdf :: Error: bad file',
            '[ERROR] 输入不存在: E:\缺失.pdf',
            'warning: 示例警告行',
            '成功：1 个文件',
            '[INFO] (1/9) Page 1',
            '[INFO] (2/9) Page 2',
            '[INFO] (3/9) Page 3',
            '[INFO] (4/9) Page 4',
            '[INFO] (5/9) Page 5',
            '[INFO] (6/9) Page 6',
            '[INFO] (7/9) Page 7',
            '[INFO] (8/9) Page 8',
            '[INFO] (9/9) Page 9',
            '[INFO] 转换完成，窗口可滚动查看全部日志。'
        )) -join "`r`n")
    }
    if ($ScreenshotSize) {
        $parts = $ScreenshotSize -split 'x'
        if ($parts.Count -eq 2) {
            $form.ClientSize = New-Object System.Drawing.Size([int]$parts[0], [int]$parts[1])
            Update-CoverBackground
        }
    }
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $form.Activate()
        Start-Sleep -Milliseconds 250
        $rect = $form.Bounds
        $bmp = New-Object System.Drawing.Bitmap($rect.Width, $rect.Height)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($rect.Location, (New-Object System.Drawing.Point(0, 0)), $rect.Size)
        $g.Dispose()
    } catch {
        $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
        $form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
    }
    $bmp.Save($ScreenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $form.Close()
    $form.Dispose()
    exit 0
}

if ($SelfTestLog) {
    $form.CreateControl()
    [void]$form.Handle
    $samples = @(
        '== CE转换引擎 转换日志 ==',
        '[OK] pdf2docx: C:\a.pdf -> C:\a.docx (1.0s)',
        '[DONE] ok=1 fail=0 outdir=C:\out',
        '[FAIL] pdf2xlsx: C:\b.pdf :: bad file',
        '[ERROR] 输入不存在: C:\c.pdf',
        '[DONE] ok=0 fail=1 outdir=C:\out',
        'Traceback (most recent call last):',
        'warning: The `fitz` API is deprecated',
        '成功：3 个文件'
    )
    foreach ($s in $samples) { [Console]::Out.WriteLine((Get-LogColorKind $s) + '|' + $s) }
    Set-LogText ($samples -join "`r`n")
    [Console]::Out.WriteLine("SELFTEST-LOG OK type=$($pnlLog.GetType().Name) chars=$((Get-LogText).Length) lines=$($script:logLines.Count)")
    $form.Dispose()
    exit 0
}

if ($SelfTestSound) {
    $form.CreateControl()
    [void]$form.Handle
    Play-ResultSound $true
    Start-Sleep -Milliseconds 700
    Play-ResultSound $false
    [Console]::Out.WriteLine('SELFTEST-SOUND OK (success + failure played)')
    $form.Dispose()
    exit 0
}

if ($SelfTest) {
    $form.CreateControl()
    [void]$form.Handle
    [Console]::Out.WriteLine("SELFTEST OK actions=$($actions.Count) files=$($lstFiles.Items.Count) python=$Py cli=$Cli")
    $form.Dispose()
    exit 0
}

[void]$form.ShowDialog()
$form.Dispose()
