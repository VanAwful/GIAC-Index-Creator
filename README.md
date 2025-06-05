# GIAC-Index-Creator
Convert your Excel spreadsheet or CSV file for the GIAC exam to a more compact and usable index. Similar to Voltaire, but using powershell and usable offline.

The work presented here was inspired by Voltaire, an on-line index application created by 
Matthew Toussain. The tool helps students create an index for the GIAC exam. It includes
the ability to paste in existing data from a spreadsheet and export the index in a CSV
format. It generates an index in a more condensed and easy to read format as a DOCX file.
https://training.opensecurity.com/

This script was created to provide an offline method to generate the same style index from a 
spreadsheet or CSV file. At the same time it allows students the flexibility to learn powershell
coding and modify it to suite their preferences.

# Requirements:
  This script was designed on Windows for Windows. 
  - Modules: All scripts require the ImportExcel module. This can be installed from the PS Gallery.
  - Excel-to-Word.ps1
      - Requires MS Word to be installed.
  - GIAC-Index-Converter.ps1 
      - Requires MS Word be installed.
  - GIAC-Index-Converter_winps.ps1 
      - Will only work with Windows PowerShell 5.1. It will not work with
        PowerShell Core. However, it does not require MS Word to be installed. 
      - Requires PSWriteWord module to be installed.


The script will read data from an MS Excel spreadsheet or CSV file and export it to an MS Word document 
in a two column index format. Currently, the required external modules can be installed from the PS
Gallery. (https://www.powershellgallery.com).

To learn more about PowerShell Gallery and how to install modules from there, visit this Microsoft site: 
 https://learn.microsoft.com/en-us/powershell/gallery/getting-started?view=powershellget-3.x

For all current scripts to work, the Excel document must have the contents in the 
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
UPDATE: The GIAC... scripts support spaces and headers. The script will present the first three lines and 
  ask the user if headers exist. If the user answers yes then the first row will be skipped. For spaces
  at the start of any value, the script now trims it as it is read in.

Here is an example of data entry (pretend it is a spreadsheet :-):

&nbsp;&nbsp;&nbsp;&nbsp;|GIAC|Global Information Assurance Certification|5|1|
   
When you are ready to create the printed index, sort the excel spreadsheet A-Z on the first column.
Then run the powershell script.
UPDATE: The GIAC... scripts will now sort the Excel spreadsheet prior to import. For CSV, you should have it sorted
  prior to running the script.

The script will format the information and output it into a Word document similar to this:
   
&nbsp;&nbsp;&nbsp;&nbsp;<b>GIAC</b> [<i>b1</i>/<i>p5</i>] Global Information Assurance Certification

The Topic will be bold and the book/page will be in italics. The description will follow and wrap as needed.
A blank line will be inserted before the next topic.

2024-08-21: Added a new script file, GIAC-Index-Converter.ps1. This builds on the Excel-to-Word including
  support added for CSV as input file.

