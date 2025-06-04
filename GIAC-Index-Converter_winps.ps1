# Required modules
Import-Module ImportExcel
Import-Module PSWriteWord

if (-not ([System.Management.Automation.PSTypeName]'System.Windows.Forms.Form').Type) {
    Add-Type -AssemblyName System.Windows.Forms
}
if (-not ([System.Management.Automation.PSTypeName]'System.Drawing.Point').Type) {
    Add-Type -AssemblyName System.Drawing
}

function Show-Form {
    try {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Document Setup"
        $form.Size = New-Object System.Drawing.Size(450, 500)
        $form.StartPosition = "CenterScreen"

        $labelInputFilePath = New-Object System.Windows.Forms.Label
        $labelInputFilePath.Text = "Input File Path:"
        $labelInputFilePath.Location = New-Object System.Drawing.Point(10, 20)
        $labelInputFilePath.Size = New-Object System.Drawing.Size(100, 20)
        $form.Controls.Add($labelInputFilePath)

        $textboxInputFilePath = New-Object System.Windows.Forms.TextBox
        $textboxInputFilePath.Size = New-Object System.Drawing.Size(250, 20)
        $textboxInputFilePath.Location = New-Object System.Drawing.Point(120, 20)
        $form.Controls.Add($textboxInputFilePath)

        $buttonBrowse = New-Object System.Windows.Forms.Button
        $buttonBrowse.Text = "Browse..."
        $buttonBrowse.Location = New-Object System.Drawing.Point(380, 18)
        $buttonBrowse.Size = New-Object System.Drawing.Size(50, 23)
        $buttonBrowse.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
            $openFileDialog.Title = "Select the Input File"
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $textboxInputFilePath.Text = $openFileDialog.FileName
                $fileExtension = [System.IO.Path]::GetExtension($openFileDialog.FileName)
                switch ($fileExtension.ToLower()) {
                    ".xlsx" { $radioExcel.Checked = $true }
                    ".csv" { $radioCSV.Checked = $true }
                }
            }
        })
        $form.Controls.Add($buttonBrowse)

        $panelInputFormat = New-Object System.Windows.Forms.Panel
        $panelInputFormat.Location = New-Object System.Drawing.Point(10, 60)
        $panelInputFormat.Size = New-Object System.Drawing.Size(200, 80)
        $form.Controls.Add($panelInputFormat)

        $labelInputFormat = New-Object System.Windows.Forms.Label
        $labelInputFormat.Text = "Input Format:"
        $labelInputFormat.Location = New-Object System.Drawing.Point(0, 0)
        $labelInputFormat.Size = New-Object System.Drawing.Size(100, 20)
        $panelInputFormat.Controls.Add($labelInputFormat)

        $radioExcel = New-Object System.Windows.Forms.RadioButton
        $radioExcel.Text = "Excel"
        $radioExcel.Location = New-Object System.Drawing.Point(10, 20)
        $radioExcel.Size = New-Object System.Drawing.Size(100, 20)
        $panelInputFormat.Controls.Add($radioExcel)

        $radioCSV = New-Object System.Windows.Forms.RadioButton
        $radioCSV.Text = "CSV"
        $radioCSV.Location = New-Object System.Drawing.Point(10, 45)
        $radioCSV.Size = New-Object System.Drawing.Size(100, 20)
        $panelInputFormat.Controls.Add($radioCSV)

        $labelMargins = New-Object System.Windows.Forms.Label
        $labelMargins.Text = "Document Margins (cm):"
        $labelMargins.Location = New-Object System.Drawing.Point(10, 160)
        $labelMargins.Size = New-Object System.Drawing.Size(200, 20)
        $form.Controls.Add($labelMargins)

        $textboxLeftMargin = New-Object System.Windows.Forms.TextBox
        $textboxLeftMargin.Text = "2.54"
        $textboxLeftMargin.Size = New-Object System.Drawing.Size(50, 20)
        $textboxLeftMargin.Location = New-Object System.Drawing.Point(120, 190)
        $form.Controls.Add($textboxLeftMargin)

        $textboxRightMargin = New-Object System.Windows.Forms.TextBox
        $textboxRightMargin.Text = "1.27"
        $textboxRightMargin.Size = New-Object System.Drawing.Size(50, 20)
        $textboxRightMargin.Location = New-Object System.Drawing.Point(120, 220)
        $form.Controls.Add($textboxRightMargin)

        $textboxTopMargin = New-Object System.Windows.Forms.TextBox
        $textboxTopMargin.Text = "0.635"
        $textboxTopMargin.Size = New-Object System.Drawing.Size(50, 20)
        $textboxTopMargin.Location = New-Object System.Drawing.Point(120, 250)
        $form.Controls.Add($textboxTopMargin)

        $textboxBottomMargin = New-Object System.Windows.Forms.TextBox
        $textboxBottomMargin.Text = "0.635"
        $textboxBottomMargin.Size = New-Object System.Drawing.Size(50, 20)
        $textboxBottomMargin.Location = New-Object System.Drawing.Point(120, 280)
        $form.Controls.Add($textboxBottomMargin)

        $buttonOK = New-Object System.Windows.Forms.Button
        $buttonOK.Text = "OK"
        $buttonOK.Location = New-Object System.Drawing.Point(175, 320)
        $buttonOK.Size = New-Object System.Drawing.Size(75, 30)
        $buttonOK.Add_Click({
            $form.Tag = @{
                InputFilePath = $textboxInputFilePath.Text
                InputFormat = if ($radioExcel.Checked) { "Excel" } else { "CSV" }
                Margins = @{
                    Left = [float]$textboxLeftMargin.Text
                    Right = [float]$textboxRightMargin.Text
                    Top = [float]$textboxTopMargin.Text
                    Bottom = [float]$textboxBottomMargin.Text
                }
            }
            $form.Close()
        })
        $form.Controls.Add($buttonOK)

        $form.ShowDialog() | Out-Null
        return $form.Tag
    } catch {
        Write-Host "Error in Show-Form: $_"
        throw
    }
}

function Get-OutputFilePath {
    try {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Word Documents (*.docx)|*.docx"
        $saveFileDialog.Title = "Save Output Word Document As"
        $saveFileDialog.FileName = "IndexOutput.docx"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $saveFileDialog.FileName
        } else {
            throw "Output path not selected. Exiting script."
        }
    } catch {
        Write-Host "Error in Get-OutputFilePath: $_"
        throw
    }
}

function Read-FirstThreeLines {
    param (
        [string]$inputFilePath,
        [string]$inputFormat
    )
    try {
        if ($inputFormat -eq "CSV") {
            $lines = Get-Content -Path $inputFilePath -TotalCount 3
        } elseif ($inputFormat -eq "Excel") {
            $data = Import-Excel -Path $inputFilePath -NoHeader -StartRow 1 -EndRow 3
            $lines = $data | ForEach-Object { $_.PSObject.Properties.Value -join ", " }
        }
        return $lines
    } catch {
        Write-Host "Error in Read-FirstThreeLines: $_"
        throw
    }
}

function Get-IfHeaders {
    param (
        [array]$lines
    )
    try {
        $headerPrompt = "The first three lines of the document are:`n"
        $headerPrompt += $lines -join "`n"
        $headerPrompt += "`n`nDoes the first line contain headers?"

        $result = [System.Windows.Forms.MessageBox]::Show($headerPrompt, "Header Detection", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        return $result
    } catch {
        Write-Host "Error in Get-IfHeaders: $_"
        throw
    }
}

function Format-DataByFirstColumn {
    param (
        [array]$data
    )
    try {
        $sortedData = $data | Sort-Object { $_.PSObject.Properties.Value[0] }
        return $sortedData
    } catch {
        Write-Host "Error in Format-DataByFirstColumn: $_"
        throw
    }
}

# Main execution block
try {
    $userInput = Show-Form
    if (-not $userInput) { throw "No input provided." }

    $inputFilePath = $userInput.InputFilePath
    $inputFormat = $userInput.InputFormat
    $margins = $userInput.Margins

    if (-not (Test-Path $inputFilePath)) {
        throw "Invalid input file path."
    }

    $outputPath = Get-OutputFilePath

    $lines = Read-FirstThreeLines -inputFilePath $inputFilePath -inputFormat $inputFormat
    $headerResponse = Get-IfHeaders -lines $lines

    if ($inputFormat -eq "CSV") {
        $data = if ($headerResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
            Import-Csv -Path $inputFilePath
        } else {
            Import-Csv -Path $inputFilePath -Header H1, H2, H3, H4
        }
    } elseif ($inputFormat -eq "Excel") {
        $data = if ($headerResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
            Import-Excel -Path $inputFilePath -NoHeader -StartRow 2
        } else {
            Import-Excel -Path $inputFilePath -NoHeader
        }
    }

    $sortedData = Format-DataByFirstColumn -data $data

    # --- PSWriteWord Section ---
    $doc = New-WordDocument -FilePath $outputPath

    # Set the section to two columns
    Set-WordSection -WordDocument $doc -Columns 2

    $previousFirstChar = ''
    $rowCount = 0   # temp counter for limiting rows

    foreach ($row in $sortedData) {
        if ($rowCount -ge 50) { break }
        $rowArray = $row.PSObject.Properties.Value
        $topic = $rowArray[0].TrimStart()
        $description = $rowArray[1]
        if ([string]::IsNullOrWhiteSpace($description)) { $description = " " }
        $page = $rowArray[2]
        $book = $rowArray[3]
        $firstChar = $topic.Substring(0, 1).ToUpper()
        if ($firstChar -notmatch '^[A-Z]$') {
            $firstChar = '#'
        }

        if ($previousFirstChar -ne $firstChar) {
            if ($previousFirstChar -ne '') {
                Add-WordPageBreak -WordDocument $doc
            }
            Add-WordText -WordDocument $doc -Text $firstChar -Bold $true -FontSize 24 -FontFamily 'Times New Roman'
            $previousFirstChar = $firstChar
        }

        # Compose the entry as an array of strings
        $bkpg = " [bk $book/pg$page] "
        $entryParts = @($topic, $bkpg, $description)

        # Apply formatting arrays: Bold for topic, Italic for bkpg, normal for description
        Add-WordText -WordDocument $doc `
            -Text $entryParts `
            -Bold $true,$false,$false `
            -Italic $false,$true,$false `
            -FontFamily 'Times New Roman','Times New Roman','Times New Roman' `
            -FontSize 10,10,10 `
            -SpacingAfter 8

        $rowCount++
    }

    Save-WordDocument -WordDocument $doc

    Write-Host "The Word document has been created successfully. Saved to $outputPath"
} catch {
    Write-Host "An error occurred: $_"
    Write-Host "Error on line: $($_.InvocationInfo.ScriptLineNumber)"
    Write-Host "Line text: $($_.InvocationInfo.Line)"
}
