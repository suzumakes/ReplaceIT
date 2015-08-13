# ReplaceIT

Using Regex on _predictable_ HTML is totally fine, and this works really well. At least for me.

First, prepare your document correctly. Save your Word document as "Web Page, Filtered (\*.htm;\*.html)".

Next, ReplaceInFolder calls ReplaceIT and CMatch so place them in a directory together and launch PowerShell.

	cd C:\Users\YourUser\YourFilePath\ReplaceIT\
	.\ReplaceInFolder.ps1 .\YourFolder--or--File

ReplaceIT will list all of the .htm files in your folder, and ask whether or not you would like to convert them.

Drop the images folder in the same directory as your converted file to make replacing your images easier.

If you're using Dreamweaver, don't forget you can click on any paragraph text in the design view and select all, cut, and paste to reformat all of your code with consistent line breaks and indentation.

###Options

	.\ReplaceInFolder.ps1 .\TargetDirectory -log

Logs all operations to a logfile in the current directory.

###Finished

* Creates a backup of every .htm it converts in the directory you specify. Need more? Edit ReplaceInFolder line 50

		$Extensions = "/.htm/"

* Cuts away almost all of M$ Word's attempts at HTML styling.
* Properly tags ~95% of super/subscripts.
* Case sensitive replacement of foreign language characters.
* Replaces all M$ Word's attempts at extracted images with a placeholder image.
* Formats tables with basic alignment, padding, and spacing.
* Exchanges p tags with bullets for list items.
* Applies a class to any internally linked words.
* Formats the resulting code into something readable and workable.

###Todo

* Even better Super/Subscript checking?

###Clone Me

If you see something that can be improved, preferably one of the todo's above, submit a pull request.

Many thanks to Michael Clark for his expertise and help. If you're here let me know so I can link to you.

