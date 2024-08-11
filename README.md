# GIAC-Index-Creator
Convert your spreadsheet for the GIAC exam to a more compact and usable index. Similar to Voltaire, but using powershell and usable offline.

The work presented here was inspired by Voltaire, an on-line index application created by 
Matthew Toussain. One function of the tool helps students convert their spreadsheet style
indexs for the GIAC exam into a more condensed and easy to read index.
https://training.opensecurity.com/

This script was created to provide an offline method to generate the same style index from a 
spreadsheet. At the same time it allows students the flexibility to learn powershell coding 
and modify it to suite their preferences.

Requirements:
  This script was designed on Windows for Windows. It does require MS Word be installed and 
  currently only reads in MS Excel files (xlsx).

The scipt will read data from an MS Excel spreadsheet and export it to an MS Word document 
in a two column index format. The only external component that is required is the 
ImportExcel module. This can be installed from the PowerShell Gallery 
(https://www.powershellgallery.com).

To learn more about PowerShell Gallery and how to install modules from there, visit this Microsoft site: 
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

&nbsp;&nbsp;&nbsp;&nbsp;|GIAC|Global Information Assurance Certification|5|1|
   
When you are ready to create the printed index, sort the excel spreadsheet A-Z on the first column.
Then run the powershell script.

The script will format the information and output it into a Word document similar to this:
   
&nbsp;&nbsp;&nbsp;&nbsp;<b>GIAC</b> [b1/p5] Global Information Assurance Certification

The Topic will be bold and the book/page will be in italics. The description will follow and wrap as needed.
A blank line will be inserted before the next topic.

