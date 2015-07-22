# ReplaceIT

There are a million reasons NOT to use Regex to parse HTML.

However, using Regex on _predicatble_ HTML is totally fine, and it oculd be your best option.

Even moreso, Word-Filtered HTML is a different story entirely. If you have to parse this garbage, these scripts will help immensely.

First, prepare your document correctly. Save your Word document as "Web Page, Filtered (\*.htm;\*.html)".

Next, ReplaceInFolder calls ReplaceIT and CMatch so place them in a directory together and launch PowerShell.

	cd C:\Users\YourUser\YourFilePath\ReplaceIT\
	.\ReplaceInFolder.ps1 .\YourFolder--or--File

ReplaceIT will list all of the .htm files in your folder, and ask whether or not you would like to convert them.

M$ tries to extract images from Word docs. I'm sure you can guess how good of a job it does. Drop this placeholder image in the same directory as your documents and you should have a much easier time replacing the images you need to.

If you're using Dreamweaver, click on any paragraph text in the design tab and select all (ctrl+a), cut (ctrl+x), and paste (ctrl+v) to reformat all of your code with consistent line breaks and indentation.

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
* Formats the resulting code into something readable and workable.

###Todo

* Format bulleted lists
* Even better Super/Subscript checking?

###Clone Me

If you see something that can be improved, preferably one of the todo's above, submit a pull request.

Many thanks to Michael Clark for his expertise and help. If you're here let me know so I can link to you.

