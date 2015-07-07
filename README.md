# ReplaceIT

There are a million reasons to NOT use Regex to parse HTML. There are very few times when this is your only option, and if it is your only option, this will help.

This is a set of PowerShell Scripts, just put ReplaceInFolder and ReplaceIT in a directory together and open PowerShell.

Run ReplaceInFolder on whathever folder contains your Word-"Filtered" HTML and it will scrape away 99% of Word's horrible useless code.

This creates a backup of every file it touches that match its acceptable file extensions.

## Want an example?

Look in convert/html/

1. 15033_1.htm.bak - This is what Word originally spat out.
2. 15033_1.htm - This is the result. Much more manageable.

Many thanks to Michael Clark for his expertise and help. If you're here let me know so I can properly credit you.