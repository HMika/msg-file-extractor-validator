# Extract attachments from .msg files via Outlook COM
# Place this script in the folder with .msg files and run it

$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir   = Join-Path $scriptDir "attachments"
$msgFiles    = Get-ChildItem -Path $scriptDir -Filter "*.msg"

if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

if ($msgFiles.Count -eq 0) {
    Write-Host "No .msg files found in: $scriptDir" -ForegroundColor Yellow
    pause
    exit
}

Write-Host "Found .msg files: $($msgFiles.Count)" -ForegroundColor Cyan
Write-Host "Output folder: $outputDir`n" -ForegroundColor Cyan

try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Host "ERROR: Could not start Outlook. Make sure Outlook is installed." -ForegroundColor Red
    pause
    exit
}

$totalExtracted = 0

foreach ($file in $msgFiles) {
    Write-Host "Processing: $($file.Name)" -ForegroundColor White

    try {
        $msg = $outlook.CreateItemFromTemplate($file.FullName)

        if ($msg.Attachments.Count -eq 0) {
            Write-Host "  [!] No attachments`n" -ForegroundColor Yellow
            $msg.Close(1)
            continue
        }

        Write-Host "  Attachments: $($msg.Attachments.Count)" -ForegroundColor Gray

        foreach ($att in $msg.Attachments) {
            $fileName = $att.FileName
            $destPath = Join-Path $outputDir $fileName

            $counter = 1
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $ext      = [System.IO.Path]::GetExtension($fileName)
            while (Test-Path $destPath) {
                $destPath = Join-Path $outputDir "$($baseName)_$counter$ext"
                $counter++
            }

            $att.SaveAsFile($destPath)
            $sizeKb = [math]::Round((Get-Item $destPath).Length / 1KB, 1)
            Write-Host "  OK  $([System.IO.Path]::GetFileName($destPath))  ($sizeKb KB)" -ForegroundColor Green
            $totalExtracted++
        }

        $msg.Close(1)

    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
    }

    Write-Host ""
}

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Done. Extracted: $totalExtracted attachment(s)" -ForegroundColor Cyan
Write-Host "Saved to: $outputDir" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

pause
