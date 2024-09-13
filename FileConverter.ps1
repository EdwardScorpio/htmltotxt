Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

# Функции конвертации (ваш код)
function Show-ProgressBar {
    param (
        [int]$PercentComplete
    )
    $width = 50
    $complete = [math]::Floor($width * ($PercentComplete / 100))
    $remaining = $width - $complete
    $progressBar = "[" + ("=" * $complete) + (" " * $remaining) + "]"
    Write-Host "`r$progressBar $PercentComplete%" -NoNewline
}

function Convert-HtmlToTxt {
    param (
        [string]$sourceFolder
    )

    $destinationFolder = Join-Path (Split-Path $sourceFolder -Parent) "$((Split-Path $sourceFolder -Leaf))_converted_txt"

    if (!(Test-Path -Path $destinationFolder)) {
        New-Item -ItemType Directory -Path $destinationFolder | Out-Null
        Write-Host "Created destination folder: $destinationFolder"
    }

    $htmlFiles = Get-ChildItem -Path $sourceFolder -Filter "*.html"
    $totalFiles = $htmlFiles.Count
    $convertedCount = 0
    $unconvertedFiles = @()
    $startTime = Get-Date

    foreach ($file in $htmlFiles) {
        $txtFileName = [System.IO.Path]::ChangeExtension($file.Name, "txt")
        $txtFilePath = Join-Path -Path $destinationFolder -ChildPath $txtFileName

        if (!(Test-Path $txtFilePath) -or (Get-Item $file).LastWriteTime -gt (Get-Item $txtFilePath).LastWriteTime) {
            $htmlContent = Get-Content -Path $file.FullName -Raw -Encoding UTF8

            $title = if ($htmlContent -match '<title>(.*?)</title>') {
                $matches[1]
            } else {
                "No title"
            }

            $preContent = if ($htmlContent -match '(?s)<pre>(.*?)</pre>') {
                $matches[1]
            } else {
                ""
            }

            $preContent = [System.Web.HttpUtility]::HtmlDecode($preContent)
            $textContent = "$title`n`n$preContent".Trim()

            $textContent | Out-File -FilePath $txtFilePath -Encoding UTF8
            $convertedCount++
        }

        $percentComplete = [math]::Round(($convertedCount / $totalFiles) * 100)
        Show-ProgressBar -PercentComplete $percentComplete
    }

    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-Host "`n`nConversion complete!"
    Write-Host "Total files processed: $totalFiles"
    Write-Host "Files converted/updated: $convertedCount"
    Write-Host "Time taken: $($duration.ToString('hh\:mm\:ss'))"
    Write-Host "Converted files are saved in: $destinationFolder"

    $unconvertedFiles = $htmlFiles | Where-Object { -not (Test-Path (Join-Path -Path $destinationFolder -ChildPath ([System.IO.Path]::ChangeExtension($_.Name, "txt")))) }

    if ($unconvertedFiles) {
        Write-Host "`nWarning: The following files were not converted:"
        $unconvertedFiles | ForEach-Object { Write-Host " - $($_.Name)" }
    }

    return $destinationFolder
}

# Основной код
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing HTML files"
$folderBrowser.ShowNewFolderButton = $false

if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $sourceFolder = $folderBrowser.SelectedPath
    $destinationFolder = Convert-HtmlToTxt -sourceFolder $sourceFolder

    $openFolder = Read-Host "`nDo you want to open the folder with converted files? (Y/N)"
    if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
        Invoke-Item $destinationFolder
    }
} else {
    Write-Host "No folder selected. Exiting."
}

# Создание формы
$form = New-Object System.Windows.Forms.Form
$form.Text = "File Converter"
$form.Size = New-Object System.Drawing.Size(400,300)
$form.StartPosition = "CenterScreen"

# Создание области для перетаскивания
$dropPanel = New-Object System.Windows.Forms.Panel
$dropPanel.Size = New-Object System.Drawing.Size(380,200)
$dropPanel.Location = New-Object System.Drawing.Point(10,10)
$dropPanel.BorderStyle = "FixedSingle"
$dropPanel.AllowDrop = $true
$form.Controls.Add($dropPanel)

# Текст в области перетаскивания
$dropLabel = New-Object System.Windows.Forms.Label
$dropLabel.Text = "Drag and drop files, folders, or ZIP archives here"
$dropLabel.AutoSize = $false
$dropLabel.Size = $dropPanel.Size
$dropLabel.TextAlign = "MiddleCenter"
$dropPanel.Controls.Add($dropLabel)

# Кнопка конвертации
$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Text = "Convert"
$convertButton.Size = New-Object System.Drawing.Size(100,30)
$convertButton.Location = New-Object System.Drawing.Point(150,230)
$convertButton.Enabled = $false
$form.Controls.Add($convertButton)

# Обработка перетаскивания
$dropPanel.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = "Copy"
    } else {
        $_.Effect = "None"
    }
})

$dropPanel.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    $dropLabel.Text = "Files added: " + $files.Count
    $convertButton.Enabled = $true
    $script:droppedFiles = $files
})

# Обработка нажатия кнопки конвертации
$convertButton.Add_Click({
    foreach ($file in $script:droppedFiles) {
        if (Test-Path $file -PathType Container) {
            # Если это папка
            Convert-HtmlToTxt -sourceFolder $file
        } elseif ($file -like "*.zip") {
            # Если это ZIP-архив
            $extractPath = [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
            [System.IO.Compression.ZipFile]::ExtractToDirectory($file, $extractPath)
            Convert-HtmlToTxt -sourceFolder $extractPath
            Remove-Item -Path $extractPath -Recurse -Force
        } else {
            # Если это отдельный файл
            $tempFolder = [System.IO.Path]::GetTempPath() + [System.IO.Path]::GetRandomFileName()
            New-Item -ItemType Directory -Path $tempFolder | Out-Null
            Copy-Item $file $tempFolder
            Convert-HtmlToTxt -sourceFolder $tempFolder
            Remove-Item -Path $tempFolder -Recurse -Force
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Conversion completed!", "Success")
})

# Показать форму
$form.ShowDialog()