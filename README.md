# msg-file-extractor-validator
# ============================================================
#  Extract-MsgAttachments.ps1
#  Извлекает вложения из .msg файлов через Outlook COM
#
#  Использование:
#    1. Положи скрипт в папку с .msg файлами
#    2. Правой кнопкой -> "Run with PowerShell"
#    3. Вложения появятся в папке "attachments" рядом
# ============================================================

# --- Настройки ---
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir   = Join-Path $scriptDir "attachments"
$msgFiles    = Get-ChildItem -Path $scriptDir -Filter "*.msg"

# --- Создать папку вывода ---
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

if ($msgFiles.Count -eq 0) {
    Write-Host "Нет .msg файлов в папке: $scriptDir" -ForegroundColor Yellow
    pause
    exit
}

Write-Host "Найдено .msg файлов: $($msgFiles.Count)" -ForegroundColor Cyan
Write-Host "Папка вывода: $outputDir`n" -ForegroundColor Cyan

# --- Запустить Outlook ---
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Host "Ошибка: не удалось запустить Outlook. Убедитесь что Outlook установлен." -ForegroundColor Red
    pause
    exit
}

$totalExtracted = 0

foreach ($file in $msgFiles) {
    Write-Host "Обрабатываю: $($file.Name)" -ForegroundColor White

    try {
        $msg = $outlook.CreateItemFromTemplate($file.FullName)

        if ($msg.Attachments.Count -eq 0) {
            Write-Host "  [!] Вложений нет`n" -ForegroundColor Yellow
            $msg.Close(1)  # 1 = olDiscard
            continue
        }

        Write-Host "  Вложений: $($msg.Attachments.Count)" -ForegroundColor Gray

        foreach ($att in $msg.Attachments) {
            $fileName = $att.FileName
            $destPath = Join-Path $outputDir $fileName

            # Если файл с таким именем уже есть — добавить счётчик
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
        Write-Host "  ОШИБКА: $_" -ForegroundColor Red
    }

    Write-Host ""
}

# --- Освободить COM ---
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Готово. Извлечено вложений: $totalExtracted" -ForegroundColor Cyan
Write-Host "Сохранено в: $outputDir" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

pause
