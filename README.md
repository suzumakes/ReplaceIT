# ReplaceIT

This cleans Word's "Web Page, Filtered" output.

###How-To

Save your Word document as "Web Page, Filtered (\*.htm;\*.html)".

Clone this repository into a local folder and place your htm files into the "convert" directory.

    cd C:\Users\Whatever\ReplaceIT
    .\ReplaceInFolder.ps1 .\convert\YourFolder-or-File

ReplaceIT will confirm the file, or list all of the .htm files in the folder you target and ask you if you would like to convert them. _(Want to submit a PR? Make it look multiple levels down)_

!! ReplaceIT _will_ search down multiple folders and convert any .htm files it finds !!

###Options

    .\ReplaceInFolder.ps1 .\TargetDirectory -log

Logs all operations to a logfile in the current directory.

###Functions

* Creates a backup before converting
* Removes inline styling
* Identifies super and subscripts (pretty damn well)
* Case-sensitive foreign language character replacement
* Replaces all images with a placeholder
* Formats tables
* Converts bulleted lists to unordered lists
* Adds "title" class to linked \<p\> elements

Many thanks to Michael Clark for his remarkable expertise.
