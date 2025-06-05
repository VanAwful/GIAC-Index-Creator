# GIAC-Index-Creator
Convert your Excel spreadsheet or CSV file for the GIAC exam to a more compact and usable index.

The work presented here was inspired by Voltaire, an on-line index application created by 
Matthew Toussain. The tool helps students create an index for the GIAC exam. It generates an index
in a more condensed and easy to read format as a DOCX file. https://training.opensecurity.com/

# History
The scripts in this project were created to provide an offline method to generate the same style index
from an Excel spreadsheet or CSV file without requiring an internet connection. This started as a 
personal project to convert my Excel based index files using PowerShell. It started with the
Excel-to-Word-index-converter.ps1. This was written on a system that had both Excel and Word installed.
It worked well and laid the ground work for the GIAC-Index-Converter.pst script. 

GIAC-Index-Converter.ps1 added more features:
  - A form to gather info from user at the start and allow margins to be set.
  - Support added to input from a CSV file.
  - Show user first few rows of data to determine if header row exist. So now you do not have to remove your header row.
  - Pre-sort the input file before processing. So user does not have to sort the input file ahead of time. (Excel)

At some point the security on my systems changed and the script would no longer auto launch MS Word. 
Rather than digging into the reason and look for a workaround, I decided it was time for a version that 
would work without requiring MS Word to be installed. This is when GIAC-Index-Converter_winps.ps1 was born.
This approach started with experimenting with various PS modules for MSWord. I landed on PSWriteWord. It took some
time, but I was able to get almost everthing working. It all looks good, but then I found PSWriteWord would not
support setting columns in the sections. So, stopped further work on it. If you want an index without the dual 
columns, give it a try. I did not get far enough to implement the section start page check (ensures each section
starts on an odd numbered page).

Feeling that I had gone as far as I could with PowerShell, I turned to a more common language, python. For this work
I turned to AI to do a quick a turn around. I used OpenAI to port my GIAC-Index-Converter.ps1 code into python. It
was not a direct port, but it did get almost 90% of it ported, which greatly reduced the time to complete it. Some of
the advantages of the GIAC-Index-Converter.py are, cross platform (no longer bound to windows), much more widely
known and used, and frankly, much easier than PowerShell, for this project. Oh, and speed. It is very fast.


# Requirements:
  - The PowerShell scripts were written to run on Windows 10/11.
  - The python scripts were written and tested with 3.12
  - All PowerShell scripts require the ImportExcel module. This can be installed from the PS Gallery.
  - Excel-to-Word ps1 equires MS Word to be installed.
  - GIAC-Index-Converter.ps1 requires MS Word be installed.
  - GIAC-Index-Converter_winps.ps1 
      - Will only work with Windows PowerShell 5.1. It will not work with
        PowerShell Core. However, it does not require MS Word to be installed. 
      - Requires PSWriteWord module to be installed.
  - GIAC-Index-Converter.py requires the following:
      - tkinter
      - pandas
      - python-docx
      - openpyxl

# Resources
For the PowerShell scripts, the required external modules can be installed from the PS Gallery. 
(https://www.powershellgallery.com).

To learn more about PowerShell Gallery and how to install modules from there, visit this Microsoft site: 
 https://learn.microsoft.com/en-us/powershell/gallery/getting-started?view=powershellget-3.x

For python, all modules were installed with pip.

# Description
The script will read data from an MS Excel spreadsheet or CSV file and export it to an MS Word document 
in a two column index format. 

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
The GIAC... scripts support spaces and headers. The script will present the first three lines and 
ask the user if headers exist. If the user answers yes then the first row will be skipped. For spaces
at the start of any value, the script now trims it as it is read in.

Here is an example of data entry (pretend it is a spreadsheet :-):

&nbsp;&nbsp;&nbsp;&nbsp;|GIAC|Global Information Assurance Certification|5|1|
   
The PowerShell GIAC... scripts will sort the Excel spreadsheet prior to import. For CSV, you should have it sorted
prior to running the script.

The script will format the information and output it into a Word document similar to this:
   
&nbsp;&nbsp;&nbsp;&nbsp;<b>GIAC</b> [<i>b1</i>/<i>p5</i>] Global Information Assurance Certification

The Topic will be bold and the book/page will be in italics. The description will follow and wrap as needed.
A blank line will be inserted before the next topic.

# Known Issues
I've stopped work on the PowerShell scripts, so will only be working on issues and features for the python
script going forward.

Issue #1: The pandas sort_values is not working. It will move entries to the end as if they have blank or
invalid values for the Topic value. This has been commented out for now. You should pre-sort the data in
Excel or CSV before importing.

Issue #2: The set 'mirror margin' is not working. For now, open the exported docx and manually set it.

