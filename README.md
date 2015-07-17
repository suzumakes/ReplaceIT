# ReplaceIT

There are a million reasons NOT to use Regex to parse HTML.

Word-Filtered HTML is a different story. If you have to parse this garbage you probably don't have a lot of other options. These scripts will help immensely.

This is a set of PowerShell Scripts, just put ReplaceInFolder and ReplaceIT in a directory together and open PowerShell.

	cd C:\Users\YourUser\YourFilePath\ReplaceIT\
	.\ReplaceInFolder.ps1 .\YourFolder--or--File

ReplaceIT will list all of the .htm files in your folder, and ask whether or not you would like to convert them.

###Done

Creates a backup of every .htm it converts in the directory you specify. Edit ReplaceInFolder line 49 to add more.

Cuts away most of M$'s Word's attempts at HTML styling, properly tags ~90% of Super/Subscripts.

Formats the resulting code into something readable and workable.

Case sensitive replacement of foreign language characters.

###Options

	.\ReplaceInFolder.ps1 .\TargetDirectory -log

Log - Logs all operations to a logfile in the current directory.

###Todo

1. Limit recursion to 1 level
2. Table Formatting
3. Bullets to Lists
4. Better Super/Subscript checking
5. Image replacement

###Clone Me

If you see something that can be improved, preferably one of the todo's above, submit a pull request.

Many thanks to Michael Clark for his expertise and help. If you're here let me know so I can link to you.

