<#
08/21/2024 by James VanOeffelen

The work presented here was inspired by Voltaire, an on-line index application created by 
Matthew Toussain. https://training.opensecurity.com/

This is a refactored version of the original Excel-to-Word-index-converter script.

Addtional functionality has been added:
  - A form to gather info from user at the start and allow temp margins to be set.
  - Support added to input CSV file, so Excel is not required.
  - Show user first few rows of data to determine if header row exist. So now you do not have to remove your header row.
  - Pre-sort the input file before processing. So user does not have to sort the input file ahead of time.
#>

# Required modules
Import-Module ImportExcel
Add-Type -AssemblyName System.Windows.Forms

# Create an MS Word Document and formats it.
function Initialize-WordDocument {
    param (
        [string]$outputPath,
        [hashtable]$margins
    )

    # Create a new Word application.
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    # Add a new document.
    $doc = $word.Documents.Add()

    # Enable Mirror Margins by setting MirrorMargins property
    $doc.PageSetup.MirrorMargins = $true

    # Set the document to two columns.
    $doc.PageSetup.TextColumns.SetCount(2)

    <# MS Word Margins
    Set up the word document for two columns and preset margins.
    Adjust these margins depending on how you will be binding the index.
    If you adjust them afterward inside Word, it could mess up the position
    of the intended Blank pages.
    To change the margins, adjust the value in centimeters.
    Common margins:
        2.54cm = 1 inch
        1.9cm = 0.75 inch
        1.27cm = 0.5 inch
        0.635 = 0.25 inch

    NOTE: If PageSetup.MirrorMargins is $true, the following margin parameters will auto-mirror
            between odd and even pages (inside and outside margin verses Left and Right).
    #>
    # Set margins for mirrored pages using the values from the form.
    $doc.PageSetup.LeftMargin = $word.CentimetersToPoints($margins.Left)  # Inside margin (binding side)
    $doc.PageSetup.RightMargin = $word.CentimetersToPoints($margins.Right) # Outside margin
    $doc.PageSetup.TopMargin = $word.CentimetersToPoints($margins.Top)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints($margins.Bottom)

    return @{
        WordApp = $word
        Document = $doc
    }
}

# Write the data to an MS Word Document.
function Add-ContentToWordDocument {
    param (
        [Microsoft.Office.Interop.Word._Document]$doc,
        [Microsoft.Office.Interop.Word.Application]$word,
        [array]$data
    )

    # Initialize the previous first character to track changes.
    # This tracks the alphabetical order: aA, bB, cC, etc.
    $previousFirstChar = ''
    $isFirstEntry = $true

    # Function to add a blank page, if needed, at the end of a section.
    function Add-PageBreakIfNeeded {
        param (
            [Microsoft.Office.Interop.Word._Document]$doc,
            [Microsoft.Office.Interop.Word.Application]$word
        )

        $currentPage = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
        if ($currentPage % 2 -ne 0) {
            $range = $doc.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

            # Collapse the range to the end of the current selection.
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

            # Insert 15 blank lines for insertion of the text BLANK near center of document.
            # Adjust '$i -lt 15' to increase or decrease the number of lines to insert.
            for ($i = 0; $i -lt 15; $i++) {
                $range.InsertParagraphAfter()
                $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            }

            # Set the paragraph alignment to right for the current page.
            $range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphRight

            # Insert the text "BLANK".
            $range.Text = "BLANK"   # If you do not want this printed on the blank page, change to an empty string: ""
            $range.Font.Name = "Arial"
            $range.Font.Size = 36   # Font size set in points.
            $range.Font.Bold = $true
            
            # Insert a paragraph after the "BLANK" text.
            $range.InsertParagraphAfter()
        }
    }

    # Process each row and add entries to the Word document
    foreach ($row in $data) {
        $rowArray = $row.PSObject.Properties.Value

        # Ensure correct mapping of columns
        $topic = $rowArray[0].TrimStart()
        $description = $rowArray[1]
        $page = $rowArray[2]
        $book = $rowArray[3]

        $firstChar = $topic.Substring(0, 1).ToUpper()

        <#
        Start a new page when the first character changes.
        This is used to ensure all sections end on an even page. When binding the index,
        it is preferred to have a section start on an odd page number for ease of 
        inserting tabbed pages or section separators. 
        If you do not want to insert the blank page, then comment out this IF statement.
        #>
        if ($previousFirstChar -ne $firstChar -and -not $isFirstEntry) {
            if (($firstChar -cmatch '[A-Z]') -or ($previousFirstChar -cmatch '[A-Z]')) {
                # Add page break if needed before starting new section
                Add-PageBreakIfNeeded -doc $doc -word $word

                $range = $doc.Content

                # Reposition cursor after the last entry.
                $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

                # Insert a page break.
                $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

            }
            $previousFirstChar = $firstChar
        }
       
        $bracketsContent = "[b$book/p$page]"
        $entry = " $description"

        # Add the topic in bold.
        $range = $doc.Content
        $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
        $range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphLeft
        $range.Text = $topic
        $range.Font.Name = "Times New Roman"
        $range.Font.Size = 10
        $range.Font.Bold = $true
        $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

        # Add the brackets content in italics.
        $range.Text = " $bracketsContent"
        $range.Font.Bold = $false
        $range.Font.Italic = $true
        $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

        # Add the description.
        $range.Text = $entry
        $range.Font.Italic = $false
        $range.InsertParagraphAfter()

        # Remove the additional paragraph mark inserted by InsertParagraphAfter.
        $range.MoveEnd([Microsoft.Office.Interop.Word.WdUnits]::wdCharacter, -1)
        $range.Text = $range.Text.TrimEnd("`r`n")

        $isFirstEntry = $false
    }

    # Final check for the last section.
    Add-PageBreakIfNeeded -doc $doc -word $word
}

# Read an Excel file.
function Read-ExcelFile {
    param (
        [string]$inputFilePath
    )

    # Read the Excel file without headers.
    $data = Import-Excel -Path $inputFilePath -NoHeader
    return $data
}

# Read a CSV file.
function Read-CSVFile {
    param (
        [string]$inputFilePath
    )

    # Read the CSV file without headers.
    $data = Import-Csv -Path $inputFilePath | ForEach-Object {
        $obj = New-Object PSObject -Property @{
            P1 = $_.Column1
            P2 = $_.Column2
            P3 = $_.Column3
            P4 = $_.Column4
        }
        $obj
    }
    return $data
}

function Read-FirstThreeLines {
    param (
        [string]$inputFilePath,
        [string]$inputFormat
    )

    if ($inputFormat -eq "CSV") {
        $lines = Get-Content -Path $inputFilePath -TotalCount 3
    } elseif ($inputFormat -eq "Excel") {
        $data = Import-Excel -Path $inputFilePath -NoHeader -StartRow 1 -EndRow 3
        $lines = $data | ForEach-Object { $_.PSObject.Properties.Value -join ", " }
    }
    
    return $lines
}

# Ensure the input file is sorted in the first column before processing.
function Sort-DataByFirstColumn {
    param (
        [array]$data
    )

    # Sort the data array by the first column
    $sortedData = $data | Sort-Object { $_.PSObject.Properties.Value[0] }
    return $sortedData
}

# Show the user the first three rows and ask if the first row is headers.
function Ask-IfHeaders {
    param (
        [array]$lines
    )

    $headerPrompt = "The first three lines of the document are:`n"
    $headerPrompt += $lines -join "`n"
    $headerPrompt += "`n`nDoes the first line contain headers?"

    $result = [System.Windows.Forms.MessageBox]::Show($headerPrompt, "Header Detection", [System.Windows.Forms.MessageBoxButtons]::YesNo)

    return $result
}

# Provide a GUI to gather basic information from user.
function Show-Form {
    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Document Setup"
    $form.Size = New-Object System.Drawing.Size(450, 500)
    $form.StartPosition = "CenterScreen"

    # Label and Textbox for Input File Path
    $labelInputFilePath = New-Object System.Windows.Forms.Label
    $labelInputFilePath.Text = "Input File Path:"
    $labelInputFilePath.Location = New-Object System.Drawing.Point(10, 20)
    $labelInputFilePath.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($labelInputFilePath)

    $textboxInputFilePath = New-Object System.Windows.Forms.TextBox
    $textboxInputFilePath.Size = New-Object System.Drawing.Size(250, 20)
    $textboxInputFilePath.Location = New-Object System.Drawing.Point(120, 20)
    $form.Controls.Add($textboxInputFilePath)

    # Button to open file dialog
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
            
            # Auto-set the Input Format based on the file extension
            $fileExtension = [System.IO.Path]::GetExtension($openFileDialog.FileName)
            switch ($fileExtension.ToLower()) {
                ".xlsx" { $radioExcel.Checked = $true }
                ".csv" { $radioCSV.Checked = $true }
            }
        }
    })
    $form.Controls.Add($buttonBrowse)

    # Panel for Input Format
    $panelInputFormat = New-Object System.Windows.Forms.Panel
    $panelInputFormat.Location = New-Object System.Drawing.Point(10, 60)
    $panelInputFormat.Size = New-Object System.Drawing.Size(200, 80)
    $form.Controls.Add($panelInputFormat)

    # Label for Input Format
    $labelInputFormat = New-Object System.Windows.Forms.Label
    $labelInputFormat.Text = "Input Format:"
    $labelInputFormat.Location = New-Object System.Drawing.Point(0, 0)
    $labelInputFormat.Size = New-Object System.Drawing.Size(100, 20)
    $panelInputFormat.Controls.Add($labelInputFormat)

    # Radio buttons for Input Format
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

    # Panel for Output Format
    $panelOutputFormat = New-Object System.Windows.Forms.Panel
    $panelOutputFormat.Location = New-Object System.Drawing.Point(10, 150)
    $panelOutputFormat.Size = New-Object System.Drawing.Size(200, 80)
    $form.Controls.Add($panelOutputFormat)

    # Label for Output Format
    $labelOutputFormat = New-Object System.Windows.Forms.Label
    $labelOutputFormat.Text = "Output Format:"
    $labelOutputFormat.Location = New-Object System.Drawing.Point(0, 0)
    $labelOutputFormat.Size = New-Object System.Drawing.Size(100, 20)
    $panelOutputFormat.Controls.Add($labelOutputFormat)

    # Radio buttons for Output Format
    $radioMSWord = New-Object System.Windows.Forms.RadioButton
    $radioMSWord.Text = "MS Word"
    $radioMSWord.Checked = $true
    $radioMSWord.Location = New-Object System.Drawing.Point(10, 20)
    $radioMSWord.Size = New-Object System.Drawing.Size(100, 20)
    $panelOutputFormat.Controls.Add($radioMSWord)

    $radioODT = New-Object System.Windows.Forms.RadioButton
    $radioODT.Text = "Open Document (ODT)"
    $radioODT.Enabled = $false
    $radioODT.Location = New-Object System.Drawing.Point(10, 45)
    $radioODT.Size = New-Object System.Drawing.Size(200, 20)
    $panelOutputFormat.Controls.Add($radioODT)

    # Subtext for ODT (disabled option)
    $labelODTNote = New-Object System.Windows.Forms.Label
    $labelODTNote.Text = "This format not yet supported."
    $labelODTNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Italic)
    $labelODTNote.ForeColor = [System.Drawing.Color]::Gray
    $labelODTNote.Location = New-Object System.Drawing.Point(140, 230)
    $labelODTNote.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($labelODTNote)

    # Label for Margins
    $labelMargins = New-Object System.Windows.Forms.Label
    $labelMargins.Text = "Document Margins (cm):"
    $labelMargins.Location = New-Object System.Drawing.Point(10, 260)
    $labelMargins.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($labelMargins)

    # Textboxes for Margins with proper spacing
    $labelLeftMargin = New-Object System.Windows.Forms.Label
    $labelLeftMargin.Text = "Left (Inside):"
    $labelLeftMargin.Location = New-Object System.Drawing.Point(10, 290)
    $labelLeftMargin.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($labelLeftMargin)

    $textboxLeftMargin = New-Object System.Windows.Forms.TextBox
    $textboxLeftMargin.Text = "2.54"
    $textboxLeftMargin.Size = New-Object System.Drawing.Size(50, 20)
    $textboxLeftMargin.Location = New-Object System.Drawing.Point(120, 290)
    $form.Controls.Add($textboxLeftMargin)

    $labelRightMargin = New-Object System.Windows.Forms.Label
    $labelRightMargin.Text = "Right (Outside):"
    $labelRightMargin.Location = New-Object System.Drawing.Point(10, 320)
    $labelRightMargin.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($labelRightMargin)

    $textboxRightMargin = New-Object System.Windows.Forms.TextBox
    $textboxRightMargin.Text = "1.27"
    $textboxRightMargin.Size = New-Object System.Drawing.Size(50, 20)
    $textboxRightMargin.Location = New-Object System.Drawing.Point(120, 320)
    $form.Controls.Add($textboxRightMargin)

    $labelTopMargin = New-Object System.Windows.Forms.Label
    $labelTopMargin.Text = "Top:"
    $labelTopMargin.Location = New-Object System.Drawing.Point(10, 350)
    $labelTopMargin.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($labelTopMargin)

    $textboxTopMargin = New-Object System.Windows.Forms.TextBox
    $textboxTopMargin.Text = "0.635"
    $textboxTopMargin.Size = New-Object System.Drawing.Size(50, 20)
    $textboxTopMargin.Location = New-Object System.Drawing.Point(120, 350)
    $form.Controls.Add($textboxTopMargin)

    $labelBottomMargin = New-Object System.Windows.Forms.Label
    $labelBottomMargin.Text = "Bottom:"
    $labelBottomMargin.Location = New-Object System.Drawing.Point(10, 380)
    $labelBottomMargin.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($labelBottomMargin)

    $textboxBottomMargin = New-Object System.Windows.Forms.TextBox
    $textboxBottomMargin.Text = "0.635"
    $textboxBottomMargin.Size = New-Object System.Drawing.Size(50, 20)
    $textboxBottomMargin.Location = New-Object System.Drawing.Point(120, 380)
    $form.Controls.Add($textboxBottomMargin)

    # OK Button
    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Location = New-Object System.Drawing.Point(175, 420)
    $buttonOK.Size = New-Object System.Drawing.Size(75, 30)
    $buttonOK.Add_Click({
        $form.Tag = @{
            InputFilePath = $textboxInputFilePath.Text
            InputFormat = if ($radioExcel.Checked) { "Excel" } else { "CSV" }
            OutputFormat = if ($radioMSWord.Checked) { "MSWord" } else { "ODT" }
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

    # Show the form and return the selections
    $form.ShowDialog() | Out-Null
    return $form.Tag
}



# Main script logic
try {
    # Show the form to get user input
    $userInput = Show-Form

    if (-not $userInput) {
        throw "No input provided. Exiting script."
    }

    # Extract user input
    $inputFilePath = $userInput.InputFilePath
    $inputFormat = $userInput.InputFormat
    $outputFormat = $userInput.OutputFormat
    $margins = $userInput.Margins

    # Validate input file path
    if (-not (Test-Path $inputFilePath)) {
        throw "Input file path is invalid or file does not exist."
    }

    # Read the first three lines to determine if there are headers
    $lines = Read-FirstThreeLines -inputFilePath $inputFilePath -inputFormat $inputFormat
    $headerResponse = Ask-IfHeaders -lines $lines

    if ($inputFormat -eq "CSV") {
        if ($headerResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Treat the first line as headers
            $data = Import-Csv -Path $inputFilePath
        } else {
            # Treat the first line as data and add default headers
            $data = Import-Csv -Path $inputFilePath -Header H1, H2, H3, H4
        }
    } elseif ($inputFormat -eq "Excel") {
        if ($headerResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
            $data = Import-Excel -Path $inputFilePath -NoHeader -StartRow 2
        } else {
            $data = Import-Excel -Path $inputFilePath -NoHeader
        }
    }

    # Sort the data by the first column
    $sortedData = Sort-DataByFirstColumn -data $data

    # Setup the Word document
    $wordDoc = Initialize-WordDocument -outputPath $outputPath -margins $margins
    Write-Host "Created new Word document."

    # Insert content into the Word document
    Add-ContentToWordDocument -doc $wordDoc.Document -word $wordDoc.WordApp -data $sortedData

    # Save the Word document
    $wordDoc.Document.SaveAs([ref] $outputPath)
    $wordDoc.Document.Close()
    $wordDoc.WordApp.Quit()

    Write-Host "The Word document has been created successfully. Saved to $outputPath"
} catch {
    Write-Host "An error occurred: $_"
    if ($wordDoc.WordApp -and $wordDoc.WordApp.Quit) {
        $wordDoc.WordApp.Quit()
    }
}
