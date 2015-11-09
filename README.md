# ReplaceIT

These Regex scripts cleans Word's "Web Page, Filtered" output into easily managed and tweaked HTML.

###How-To

Save your Word document as "Web Page, Filtered (\*.htm;\*.html)".

Clone this repository into a local folder and place your htm files into the "convert" directory.

CD into the folder:

	.\ReplaceInFolder.ps1 .\YourFolder\YourFile(optional)

ReplaceIT will confirm the file, or list all of the .htm files in the folder you target and ask you if you would like to convert them.

!! ReplaceIT _will_ search down multiple folder leves and convert any .htm files it finds !!

###Options

	.\ReplaceInFolder.ps1 .\TargetDirectory -log

Logs all operations to a logfile in the current directory.

###Functions

* Creates backup files before converting
* Removes Word's inline styling
* Identifies super and subscripts (mostly)
* Replaces Foreign Language Characters with ASCII equivalients (case-sensitive)
* Replaces all images with a placeholder
* Basic table formatting
* Identifies bulleted lists as \<li\> items (you will need to add \<ul\>|\<ol\> tags)
* Adds "title" class to linked \<p\> elements

###Todo

* Suggestions?

###Clone Me

If you see something that can be improved, preferably one of the todo's above, submit a pull request.

Many thanks to Michael Clark for his remarkable expertise.
