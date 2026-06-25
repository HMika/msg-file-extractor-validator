# Extract attachments from .msg files via Outlook COM
# Creates a subfolder per .msg named after the version (e.g. v5.8.0)
# Then compares files across version folders using fc and writes report

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

        $outputDir = Join-Path $scriptDir $version
        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
            Write-Host "  Created folder: $outputDir" -ForegroundColor Gray
        } else {
            Write-Host "  Using existing folder: $outputDir" -ForegroundColor Gray
        }

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
Write-Host "============================================`n" -ForegroundColor Cyan

# -------------------------------------------------------
# COMPARE FILES ACROSS VERSION FOLDERS
# -------------------------------------------------------
Write-Host "Starting file comparison..." -ForegroundColor Cyan

$versionFolders = Get-ChildItem -Path $scriptDir -Directory |
    Where-Object { $_.Name -match '^v\d+\.\d+\.\d+$' } |
    Sort-Object Name

$reportPath = Join-Path $scriptDir "comparison_report.txt"
$report     = [System.Collections.Generic.List[string]]::new()

$report.Add("FILE COMPARISON REPORT")
$report.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$report.Add("Folders compared: $($versionFolders.Name -join ', ')")
$report.Add("=" * 60)
$report.Add("")

if ($versionFolders.Count -lt 2) {
    $report.Add("ERROR: Need at least 2 version folders to compare.")
    $report | Set-Content -Path $reportPath -Encoding UTF8
    Write-Host "Not enough version folders to compare (need at least 2)." -ForegroundColor Yellow
} else {
    # Build map: baseName -> { version -> fullPath }
    $allFiles = @{}

    foreach ($folder in $versionFolders) {
        $files = Get-ChildItem -Path $folder.FullName -File
        foreach ($f in $files) {
            $ext       = $f.Extension
            $nameNoExt = $f.BaseName
            $baseName  = $nameNoExt -replace '_v\d+\.\d+\.\d+$', ''

            if (-not $allFiles.ContainsKey($baseName)) {
                $allFiles[$baseName] = @{}
            }
            $allFiles[$baseName][$folder.Name] = $f.FullName
        }
    }

    $folderList = $versionFolders.Name

    for ($i = 0; $i -lt $folderList.Count - 1; $i++) {
        for ($j = $i + 1; $j -lt $folderList.Count; $j++) {
            $verA = $folderList[$i]
            $verB = $folderList[$j]

            $report.Add("COMPARING: $verA  vs  $verB")
            $report.Add("-" * 60)

            foreach ($baseName in ($allFiles.Keys | Sort-Object)) {
                $pathA = $allFiles[$baseName][$verA]
                $pathB = $allFiles[$baseName][$verB]

                if (-not $pathA) {
                    $report.Add("  [MISSING in $verA]  $baseName")
                    continue
                }
                if (-not $pathB) {
                    $report.Add("  [MISSING in $verB]  $baseName")
                    continue
                }

                $report.Add("")
                $report.Add("  FILE: $baseName")
                $report.Add("    $verA : $pathA")
                $report.Add("    $verB : $pathB")
                $report.Add("")

                # Run fc via cmd /c to avoid PowerShell parsing issues
                $fcOutput = cmd /c "fc /L `"$pathA`" `"$pathB`"" 2>&1

                foreach ($line in $fcOutput) {
                    $report.Add("    $line")
                }
                $report.Add("")
            }

            $report.Add("")
        }
    }

    $report | Set-Content -Path $reportPath -Encoding UTF8
    Write-Host "Comparison report saved to: $reportPath" -ForegroundColor Green
}

pause
