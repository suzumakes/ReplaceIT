# ReplaceIT

There are a million reasons NOT to use Regex to parse HTML.

Word-Filtered HTML is a different story. If you have to parse this garbage you probably don't have a lot of other options. These scripts help immensely.

ReplaceInFolder calls ReplaceIT and CMatch so place them in a directory together and launch PowerShell.

	cd C:\Users\YourUser\YourFilePath\ReplaceIT\
	.\ReplaceInFolder.ps1 .\YourFolder--or--File

ReplaceIT will list all of the .htm files in your folder, and ask whether or not you would like to convert them.

M$ tries to extract images from Word docs. I'm sure you can guess how good of a job it does. Drop this placeholder image in the same directory as your documents and you should have a much easier time replacing the images you need to.

###Options

	.\ReplaceInFolder.ps1 .\TargetDirectory -log

Logs all operations to a logfile in the current directory.

###Finished

* Creates a backup of every .htm it converts in the directory you specify. Need more? Edit ReplaceInFolder line 50

		$Extensions = "/.htm/"

* Cuts away most of M$'s Word's attempts at HTML styling, properly tags ~90% of Super/Subscripts.
* Case sensitive replacement of foreign language characters.
* Replaces all M$ Word's extracted image references with a small placeholder.
* Formats the resulting code into something readable and workable.

###Todo

* Limit recursion to 1 level
* Table Formatting
* Bullets to Lists
* Better Super/Subscript checking

###Clone Me

If you see something that can be improved, preferably one of the todo's above, submit a pull request.

Many thanks to Michael Clark for his expertise and help. If you're here let me know so I can link to you.

