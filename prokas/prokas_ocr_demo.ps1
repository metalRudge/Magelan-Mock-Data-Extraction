<#
.SYNOPSIS
    Mini-PROKAS OCR Reader: Captures the PROKAS window, extracts text & fields
    via built-in Windows OCR, and stays running in the background to rescan
    the product list when signaled.
#>
[CmdletBinding()]
param(
    [string]$WindowTitle = 'PROKAS Kasse (Mock)',
    [string]$Language    = 'de-DE',
    [ValidateRange(1, 4)][double]$Scale = 2,
    [string]$OutDir,
    [switch]$LaunchMock,
    [switch]$OneShot,
    [switch]$Json
)

$ErrorActionPreference = 'Stop'
if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot 'out' }
if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'Please run with Windows PowerShell 5.1 (powershell.exe), not pwsh.'
}

# ===========================================================================
# 1. Win32 & WinRT Interop Setup
# ===========================================================================
Add-Type -AssemblyName System.Drawing, System.Runtime.WindowsRuntime
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ProkasWin32 {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    public static IntPtr FindLargestChildControl(IntPtr hWndParent) {
        IntPtr best = IntPtr.Zero;
        int bestArea = 0;
        EnumChildWindows(hWndParent, (child, _) => {
            RECT r;
            if (GetWindowRect(child, out r)) {
                int area = (r.Right - r.Left) * (r.Bottom - r.Top);
                if (area > bestArea) { bestArea = area; best = child; }
            }
            return true;
        }, IntPtr.Zero);
        return best;
    }
}
'@
[void][ProkasWin32]::SetProcessDPIAware()

$null = [Windows.Storage.StorageFile,                Windows.Storage,         ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder,     Windows.Graphics,        ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,                Windows.Foundation,      ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language,             Windows.Globalization,   ContentType = WindowsRuntime]

$script:AsTaskMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1

function Wait-WinRtTask($Operation, [Type]$ResultType) {
    $task = $script:AsTaskMethod.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    [void]$task.Wait(-1)
    return $task.Result
}

function Write-Info([string]$Message) {
    if (-not $Json) { Write-Host $Message -ForegroundColor Blue }
}

# ===========================================================================
# 2. Window Discovery & Capture Helpers
# ===========================================================================
function Find-TargetWindow([string]$TitlePart) {
    Get-Process |
        Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle.Contains($TitlePart) } |
        Select-Object -First 1
}

function Get-TargetProcess([string]$TitlePart, [bool]$ShouldLaunch) {
    $proc = Find-TargetWindow $TitlePart
    if (-not $proc -and $ShouldLaunch) {
        Write-Info 'Starting mock PROKAS window...'
        $mockScript = Join-Path $PSScriptRoot 'mock_prokas.ps1'
        [void](Start-Process powershell.exe -PassThru -ArgumentList @(
            '-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$mockScript`""
        ))
        $deadline = (Get-Date).AddSeconds(15)
        while (-not $proc -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 300
            $proc = Find-TargetWindow $TitlePart
        }
        Start-Sleep -Milliseconds 700
    }
    if (-not $proc) {
        throw "No window found with title like '*$TitlePart*'. Use -LaunchMock."
    }
    return $proc
}

function Get-BitmapSha256([System.Drawing.Bitmap]$Bitmap) {
    $stream = New-Object System.IO.MemoryStream
    $Bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
    $hashBytes = $script:Sha256.ComputeHash($stream.ToArray())
    $stream.Dispose()
    return ([BitConverter]::ToString($hashBytes) -replace '-', '')
}

function Get-WindowCapture([IntPtr]$WindowHandle, [bool]$ProductListOnly) {
    if ([ProkasWin32]::IsIconic($WindowHandle)) {
        [void][ProkasWin32]::ShowWindow($WindowHandle, 9)
        Start-Sleep -Milliseconds 300
    }

    $winRect = New-Object ProkasWin32+RECT
    [void][ProkasWin32]::GetWindowRect($WindowHandle, [ref]$winRect)
    $winW = $winRect.Right - $winRect.Left
    $winH = $winRect.Bottom - $winRect.Top

    $fullBmp = New-Object System.Drawing.Bitmap($winW, $winH)
    $g = [System.Drawing.Graphics]::FromImage($fullBmp)
    $hdc = $g.GetHdc()
    $ok = [ProkasWin32]::PrintWindow($WindowHandle, $hdc, 2)
    $g.ReleaseHdc($hdc)
    $g.Dispose()
    if (-not $ok) { $fullBmp.Dispose(); throw 'PrintWindow failed.' }

    # Locate the product ListBox control inside the window
    $listHwnd = [ProkasWin32]::FindLargestChildControl($WindowHandle)
    $cropRect = New-Object System.Drawing.Rectangle(0, 0, $winW, $winH)
    $listHash = $null

    if ($listHwnd -ne [IntPtr]::Zero) {
        $listRect = New-Object ProkasWin32+RECT
        [void][ProkasWin32]::GetWindowRect($listHwnd, [ref]$listRect)
        $lx = [Math]::Max(0, $listRect.Left - $winRect.Left)
        $ly = [Math]::Max(0, $listRect.Top  - $winRect.Top)
        $lw = [Math]::Min($winW - $lx, $listRect.Right  - $listRect.Left)
        $lh = [Math]::Min($winH - $ly, $listRect.Bottom - $listRect.Top)

        $listBmp  = $fullBmp.Clone((New-Object System.Drawing.Rectangle($lx, $ly, $lw, $lh)), $fullBmp.PixelFormat)
        $listHash = Get-BitmapSha256 $listBmp
        $listBmp.Dispose()

        if ($ProductListOnly) {
            $cropRect = New-Object System.Drawing.Rectangle($lx, $ly, $lw, ($winH - $ly))
        }
    } else {
        $listHash = Get-BitmapSha256 $fullBmp
    }

    $targetBmp = if ($ProductListOnly -and $cropRect.Y -gt 0) {
        $cropped = $fullBmp.Clone($cropRect, $fullBmp.PixelFormat)
        $fullBmp.Dispose()
        $cropped
    } else {
        $fullBmp
    }

    return [pscustomobject]@{
        Bitmap     = $targetBmp
        ListHash   = $listHash
        WindowRect = $winRect
        WinWidth   = $winW
        WinHeight  = $winH
        CropRect   = $cropRect
    }
}

# ===========================================================================
# 3. OCR Execution & Business Field Parsing
# ===========================================================================
function Invoke-OcrOnCapture($Capture, [string]$PngPath) {
    $bmp = $Capture.Bitmap
    $scaledW = [int]($bmp.Width  * $Scale)
    $scaledH = [int]($bmp.Height * $Scale)

    $scaledBmp = New-Object System.Drawing.Bitmap($scaledW, $scaledH)
    $g = [System.Drawing.Graphics]::FromImage($scaledBmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($bmp, 0, 0, $scaledW, $scaledH)
    $g.Dispose()

    $bmp.Save($PngPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $tempOcrPath = Join-Path $env:TEMP ("prokas_ocr_{0}.png" -f [guid]::NewGuid().ToString('N'))
    $scaledBmp.Save($tempOcrPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $scaledBmp.Dispose()

    $file   = Wait-WinRtTask ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tempOcrPath)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinRtTask ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder   = Wait-WinRtTask ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $softBmp   = Wait-WinRtTask ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $ocrResult = Wait-WinRtTask ($script:OcrEngine.RecognizeAsync($softBmp)) ([Windows.Media.Ocr.OcrResult])
    } finally {
        $stream.Dispose()
        Remove-Item $tempOcrPath -ErrorAction SilentlyContinue
    }

    $cropX = $Capture.CropRect.X
    $cropY = $Capture.CropRect.Y

    foreach ($line in $ocrResult.Lines) {
        $xs = $line.Words | ForEach-Object { $_.BoundingRect.X }
        $ys = $line.Words | ForEach-Object { $_.BoundingRect.Y }
        $xe = $line.Words | ForEach-Object { $_.BoundingRect.X + $_.BoundingRect.Width }
        $ye = $line.Words | ForEach-Object { $_.BoundingRect.Y + $_.BoundingRect.Height }

        $localX = [int](($xs | Measure-Object -Minimum).Minimum / $Scale)
        $localY = [int](($ys | Measure-Object -Minimum).Minimum / $Scale)
        $w      = [int](($xe | Measure-Object -Maximum).Maximum / $Scale) - $localX
        $h      = [int](($ye | Measure-Object -Maximum).Maximum / $Scale) - $localY
        $winX   = $cropX + $localX
        $winY   = $cropY + $localY

        [pscustomobject]@{
            text    = $line.Text
            x       = $winX
            y       = $winY
            w       = $w
            h       = $h
            relX    = [math]::Round(($winX + $w / 2.0) / $Capture.WinWidth,  3)
            relY    = [math]::Round(($winY + $h / 2.0) / $Capture.WinHeight, 3)
            screenX = $Capture.WindowRect.Left + $winX + [int]($w / 2)
            screenY = $Capture.WindowRect.Top  + $winY + [int]($h / 2)
        }
    }
}

function Test-PznChecksum([string]$Pzn) {
    if ($Pzn -notmatch '^\d{8}$') { return $false }
    $sum = 0
    for ($i = 0; $i -lt 7; $i++) {
        $sum += [int]::Parse($Pzn[$i]) * ($i + 1)
    }
    $checkDigit = $sum % 11
    return ($checkDigit -ne 10 -and $checkDigit -eq [int]::Parse($Pzn[7]))
}

function ConvertTo-ParsedFields([array]$Lines) {
    $fullText  = ($Lines | ForEach-Object text) -join "`n"
    $zeroChars = "Oo$([char]0x00D8)$([char]0x00F8)"
    $digitText = [regex]::Replace($fullText, '(?<=\d)[Oo\u00D8\u00F8Il]|[Oo\u00D8\u00F8Il](?=\d)', {
        param($m) if ($zeroChars.Contains($m.Value)) { '0' } else { '1' }
    })

    $status = if ($fullText -match 'Status:?\s*([\p{L} ]+)') { $Matches[1].Trim() }
              elseif ($script:CurrentState) { $script:CurrentState.fields.status }
              else { 'ERFASST' }

    $total  = if ($digitText -match 'Total:?\s*CHF\s*(\d+[.,]\d{2})') { [double]($Matches[1] -replace ',', '.') }
              elseif ($script:CurrentState) { $script:CurrentState.fields.totalChf }
              else { $null }

    $pzns = [regex]::Matches($digitText, '(?<!\d)\d{8}(?!\d)') |
        ForEach-Object Value |
        Select-Object -Unique |
        ForEach-Object { [pscustomobject]@{ pzn = $_; checksumValid = (Test-PznChecksum $_) } }

    return [ordered]@{
        status   = $status
        totalChf = $total
        pzns     = @($pzns)
    }
}

# ===========================================================================
# 4. Scan Orchestrator & Reporting
# ===========================================================================
function Write-ScanReport($State, [array]$ScannedLines, $CropRect, [string]$LatestJsonPath, [bool]$ProductListOnly) {
    if ($Json) {
        [ordered]@{
            window   = $State.window
            scanMode = $State.scanMode
            status   = $State.fields.status
            totalChf = $State.fields.totalChf
            pzns     = $State.fields.pzns
            lines    = @($ScannedLines | ForEach-Object text)
            detail   = $LatestJsonPath
        } | ConvertTo-Json -Depth 4 -Compress
        return
    }

    $modeLabel = if ($ProductListOnly) { 'PRODUCT LIST RESCAN' } else { 'INITIAL FULL SCAN' }
    Write-Host "`n=== [$modeLabel] OCR lines ($($ScannedLines.Count)) | Crop: $($CropRect.Width)x$($CropRect.Height)px ===" -ForegroundColor Yellow
    $ScannedLines | Format-Table text, x, y, w, h, relX, relY, screenX, screenY -AutoSize | Out-String | Write-Host

    Write-Host '=== Parsed fields ===' -ForegroundColor Yellow
    Write-Host ("Status : {0}" -f $State.fields.status)
    Write-Host ("Total  : CHF {0:N2}" -f $State.fields.totalChf)
    foreach ($p in $State.fields.pzns) {
        $statusText = if ($p.checksumValid) { 'OK' } else { 'INVALID' }
        Write-Host ("PZN    : {0}  checksum {1}" -f $p.pzn, $statusText)
    }
    Write-Host "Updated JSON -> $LatestJsonPath" -ForegroundColor Green
}

function Invoke-ProkasScan([switch]$ProductListOnly, [switch]$ForceExtract) {
    $capture = Get-WindowCapture $script:TargetHwnd $ProductListOnly

    $unchanged = ($ProductListOnly -and (-not $ForceExtract) -and ($capture.ListHash -eq $script:LastListHash))
    if ($unchanged) {
        $capture.Bitmap.Dispose()
        Write-Info "[$(Get-Date -Format 'HH:mm:ss')] No change in product list - skipping OCR."
        return
    }
    $script:LastListHash = $capture.ListHash

    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $tag     = if ($ProductListOnly) { 'rescan' } else { 'init' }
    $pngPath = Join-Path $OutDir "prokas_${tag}_$stamp.png"

    $scannedLines = @(Invoke-OcrOnCapture $capture $pngPath)
    $parsedFields = ConvertTo-ParsedFields $scannedLines

    $mergedLines = if ($ProductListOnly -and $script:CurrentState) {
        $headerLines = $script:CurrentState.lines | Where-Object { $_.y -lt $capture.CropRect.Y }
        @($headerLines) + $scannedLines
    } else {
        $scannedLines
    }

    $script:CurrentState = [ordered]@{
        window     = $script:TargetProcess.MainWindowTitle
        scanMode   = if ($ProductListOnly) { 'ProductListRescan' } else { 'InitialFullScan' }
        capturedAt = (Get-Date).ToString('s')
        screenshot = $pngPath
        windowRect = [ordered]@{
            left   = $capture.WindowRect.Left
            top    = $capture.WindowRect.Top
            width  = $capture.WinWidth
            height = $capture.WinHeight
        }
        fields     = $parsedFields
        lines      = $mergedLines
    }

    $snapshotJsonPath = Join-Path $OutDir "prokas_$stamp.json"
    $latestJsonPath   = Join-Path $OutDir 'prokas_latest.json'
    $jsonContent      = $script:CurrentState | ConvertTo-Json -Depth 5
    $jsonContent | Set-Content -Path $snapshotJsonPath -Encoding UTF8
    $jsonContent | Set-Content -Path $latestJsonPath   -Encoding UTF8

    Write-ScanReport $script:CurrentState $scannedLines $capture.CropRect $latestJsonPath $ProductListOnly
}





New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$langObj = New-Object Windows.Globalization.Language($Language)
if (-not [Windows.Media.Ocr.OcrEngine]::IsLanguageSupported($langObj)) {
    throw "OCR language '$Language' is not installed on this machine."
}

$script:OcrEngine     = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($langObj)
$script:Sha256        = [System.Security.Cryptography.SHA256]::Create()
$script:TargetProcess = Get-TargetProcess $WindowTitle $LaunchMock.IsPresent
$script:TargetHwnd    = $script:TargetProcess.MainWindowHandle
$script:LastListHash  = $null
$script:CurrentState  = $null

$rescanEvent       = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', 'Local\ProkasOcrRescan')
$manualRescanEvent = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', 'Local\ProkasOcrManualRescan')
[void]$rescanEvent.Reset()
[void]$manualRescanEvent.Reset()

# Run baseline full scan
Invoke-ProkasScan

if ($OneShot) { return }

Write-Info "`n[OCR Engine Running in Background] Waiting for new product or 'OCR Rescan' button..."
while ([ProkasWin32]::IsWindow($script:TargetHwnd)) {
    if ($rescanEvent.WaitOne(500)) {
        $isManual = $manualRescanEvent.WaitOne(0)
        Start-Sleep -Milliseconds 150
        [void]$rescanEvent.Reset()
        [void]$manualRescanEvent.Reset()

        if ([ProkasWin32]::IsWindow($script:TargetHwnd)) {
            Invoke-ProkasScan -ProductListOnly -ForceExtract:$isManual
        }
    }
}
Write-Info 'PROKAS window closed. Background OCR engine stopped.'
