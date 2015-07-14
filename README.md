# ReplaceIT

There are a million reasons NOT to use Regex to parse Word-Filtered HTML. There are some times when this is your only option, and this will help.

This is a set of PowerShell Scripts, just put ReplaceInFolder and ReplaceIT in a directory together and open PowerShell.

	cd C:\Users\YourUser\YourFilePath\ReplaceIT
	.\ReplaceInFolder.ps1 .\YourFolder

ReplaceIT will list all of the .htm files in your folder, and ask whether or not you would like to convert them. (You can just hit enter to accept)

###Options

Remember in PowerShell the options go after the target directory or file.

	.\ReplaceInFolder.ps1 .\TargetDirectory -vb

Verbose - The default is pretty quiet about what's happening to your files, verbose will fill your stdout with messages about everything it can and can't find.

### What it Does

Creates a backup (.bak) of every .htm file in the directory you specify. You can edit the acceptable extensions to add more if you like.

Creates a log file of all the changes it makes, alter the directory on line 77 of ReplaceIT.ps1 to put the logs where you like. 

Scrapes away 99% of Word's horrible useless code, leaving you with more free time and fewer headaches.

It's pretty good at formatting the resulting code into something workable too.

###Todo

1. Limit recursion to 1 level
2. Table Formatting
3. Bullets to Lists
4. Better super and subscript checking
5. Strict case sensitivity for Foreign Language Characters

###Want to Contribute?

Fork ReplaceIT and submit a pull request if you want to contribute! The Todo's are next in development but if you see something that can be improved please let me know!

It would make me really happy to know someone else benefits from this, so if it helps you out let me know!

Many thanks to Michael Clark for his expertise and help. If you're here let me know so I can properly credit you. Seriously my man.

