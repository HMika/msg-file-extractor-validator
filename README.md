# Extract attachments from .msg files via Outlook COM
# Creates a subfolder per .msg named after the version (e.g. v5.8.0)
# Renames each attachment with the version suffix

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$msgFiles   = Get-ChildItem -Path $scriptDir -Filter "*.msg"

if ($msgFiles.Count -eq 0) {
    Write-Host "No .msg files found in: $scriptDir" -ForegroundColor Yellow
    pause
    exit
}

Write-Host "Found .msg files: $($msgFiles.Count)" -ForegroundColor Cyan

try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Host "ERROR: Could not start Outlook." -ForegroundColor Red
    pause
    exit
}

$totalExtracted = 0

foreach ($file in $msgFiles) {
    Write-Host "`nProcessing: $($file.Name)" -ForegroundColor White

    try {
        $msg = $outlook.CreateItemFromTemplate($file.FullName)

        # Find version vX.Y.Z in subject or filename
        $version = $null

        if ($msg.Subject -match 'v(\d+\.\d+\.\d+)') {
            $version = "v$($Matches[1])"
            Write-Host "  Version found in subject: $version" -ForegroundColor DarkCyan
        }
        elseif ($file.Name -match 'v(\d+\.\d+\.\d+)') {
            $version = "v$($Matches[1])"
            Write-Host "  Version found in filename: $version" -ForegroundColor DarkCyan
        }
        else {
            $version = "unknown_version"
            Write-Host "  [!] No version found - folder will be named 'unknown_version'" -ForegroundColor Yellow
        }

        # Create subfolder named after version
        $outputDir = Join-Path $scriptDir $version
        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
        }
        Write-Host "  Output folder: $outputDir" -ForegroundColor Gray

        if ($msg.Attachments.Count -eq 0) {
            Write-Host "  [!] No attachments" -ForegroundColor Yellow
            $msg.Close(1)
            continue
        }

        Write-Host "  Attachments: $($msg.Attachments.Count)" -ForegroundColor Gray

        foreach ($att in $msg.Attachments) {
            $origName = $att.FileName
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($origName)
            $ext      = [System.IO.Path]::GetExtension($origName)

            $newName  = "$($baseName)_$($version)$ext"
            $destPath = Join-Path $outputDir $newName

            # Handle duplicates
            $counter = 1
            while (Test-Path $destPath) {
                $destPath = Join-Path $outputDir "$($baseName)_$($version)_$counter$ext"
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
}

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "Done. Extracted: $totalExtracted attachment(s)" -ForegroundColor Cyan
Write-Host "Folders created in: $scriptDir" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

pause
