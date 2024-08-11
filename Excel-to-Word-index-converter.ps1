<# Excel to Word index converter.
 Created 07-2024 by James VanOeffelen, GICSP, GCWN, GLSC, GPYC, GAWN

 The work presented here was inspired by Voltaire, an on-line index application created by 
 Matthew Toussain. It was created as tool to help students convert their spreadsheet style
 indexs for the GIAC exam into a more condensed and easy to read index.
 https://training.opensecurity.com/

 This script was created to provide an offline method to generate an index from a spreadsheet. At
 the same time it allows students the flexibility to learn the powershell coding and modify it to
 suite their preferences.
 
 Requirements:
    This script was designed on Windows for Windows. It does require MS Word be installed and 
    currently only reads in MS Excel files (xlsx).

 The scipt will read data from an MS Excel spreadsheet and export it to a Microsoft Word
 document in a two column index format. The only external component that is required is the 
 ImportExcel module. This can be installed from the PowerShell Gallery 
 (https://www.powershellgallery.com). To learn more about PowerShell Gallery and how to install
 modules from there, visit this Microsoft site: 
   https://learn.microsoft.com/en-us/powershell/gallery/getting-started?view=powershellget-3.x

 For the current version of the script to work, the Excel document must have the contents in the 
 following format:
    - Four columns are used.
    - First column is the Topic. This is the word or words that the entire index will be sorted on.
    - Second column is the Description. This is useful to provide brief information about the
      topic. For example, a definition. You as much or as little in the description as you want. Just
      note that too much information may slow down your search for the answer.
    - Third column is Page number it is found on.
    - Fourth column is the book number it is found in. 
 The contents of the cells must not start with a space and no empty rows between entries.
 Remove the headers if they exist as they are not yet supported.
 Here is an example of data entry (pretend it is a spreadsheet :-):
     -----------------------------------------------------
     |GIAC|Global Information Assurance Certification|5|1|
     -----------------------------------------------------

When you are ready to create the printed index, sort the excel spreadsheet A-Z on the first column.
Then run the powershell script.

 The script will format the information and output it into a Word document similar to this:
     GIAC [b1/p5] Global Information
     Assurance Certification

 The Topic will be bold and the book/page will be in italics. The description will follow and wrap as needed.
 A blank line will be inserted before the next topic.

*** Versioning ***
1.2.2: Provided vertical and horitzonal alignment of BLANK text on blank page. Added text format and more instructions.
1.2.1: Corrected issue with cursor insertion point moving to the start of the document during BLANK insertion.
1.2.0: Added blank page insertion to make each section end on an even page. Helps with section dividers when printed.
1.1.0: Added margin control for new word doc. Set Topic and book/page font styles.
1.0.0: Initial release with alphabetically sorted index topics in two columns. 

#>

# Required modules.
Import-Module ImportExcel

# Required assemblies for GUI.
Add-Type -AssemblyName System.Windows.Forms

try {
    # Create an OpenFileDialog object.
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $openFileDialog.Title = "Select the Excel file"

    # Show the file dialog and get the selected file.
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $excelPath = $openFileDialog.FileName
    } else {
        throw "No file selected. Exiting script."
    }
    
    # Extract directory and file name without extension.
    $excelDir = Split-Path -Path $excelPath -Parent
    $excelName = [System.IO.Path]::GetFileNameWithoutExtension($excelPath)
    $outputPath = Join-Path -Path $excelDir -ChildPath "$excelName.docx"

    # Read the Excel file without headers.
    $data = Import-Excel -Path $excelPath -NoHeader

    # Create a new Word application.
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    # Add a new document.
    $doc = $word.Documents.Add()

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
    #>
    $doc.PageSetup.TextColumns.SetCount(2)
    $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(2.54)
    $doc.PageSetup.RightMargin = $word.CentimetersToPoints(1.27)
    $doc.PageSetup.TopMargin = $word.CentimetersToPoints(0.635)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(0.635)

    # Initialize the previous first character to track changes.
    # This tracks the alphabetical order: aA, bB, cC, etc.
    $previousFirstChar = ''

    # Function to add a blank page, if needed, at the end of a section.
    function Add-PageBreakIfNeeded {
        param (
            [Microsoft.Office.Interop.Word.Document]$doc,
            [Microsoft.Office.Interop.Word.Application]$word
        )

        $currentPage = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
        if ($currentPage % 2 -ne 0) {
            $range = $doc.Content
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

            # Collapse the range to the end of the current selection.
            $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

            # Insert 20 blank lines for insertion of the text BLANK near center of document.
            # Adjust '$i -lt 20' to increase or decrease the number of lines to insert.
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
    $isFirstEntry = $true
    foreach ($row in $data) {
        $rowArray = $row.PSObject.Properties.Value

        # Ensure correct mapping of columns
        $topic = $rowArray[0].TrimStart()
        $description = $rowArray[1]
        $page = $rowArray[2]
        $book = $rowArray[3]

        $firstChar = $topic.Substring(0, 1).ToUpper()

        ############################
        # Start a new page when the first character changes.
        # This is used to ensure all sections end on an even page. When binding the index,
        # it is preferred to have a section start on an odd page number for ease of 
        # inserting tabbed pages or section separators. 
        # If you do not want to insert the blank page, then comment out this IF statement.
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
        ############################

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

    # Final check for last section.
    Add-PageBreakIfNeeded -doc $doc -word $word

    # Save the Word document.
    $doc.SaveAs([ref] $outputPath)
    $doc.Close()
    $word.Quit()

    Write-Host "The Word document has been created successfully. Saved to $outputPath"
} catch {
    Write-Host "An error occurred: $_"
    if ($word -and $word.Quit) {
        $word.Quit()
    }
}
