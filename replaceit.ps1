<#

This is a Find and Replace PowerShell Script for cleaning Word-Generated HTML.

Many thanks to Michael Clark for the inital replaceit PowerShell Script.

====================
TO DO

1. Implement user confirmation - i.e. "would you like to run ReplaceIT on file_a?
2. Restrict editing to .txt .htm and .html
3. Solicit help from any PowerShell Experts to condense/improve codebase.

====================

#>

Param (
	[string]$File,
	[string]$Find,
	[string]$Replace,
	[switch]$AllMatches,
	[string]$Pattern,
	[string]$Start,
	[string]$End,
	[string]$NewStart,
	[string]$NewEnd
)
Function SuperScript
	{
		Param (
			$FileContents,
			$Pattern,
			$Start,
			$End,
			$NewStart,
			$NewEnd,
			$File
		)

		$MyMatches = $FileContents | Select-String -Pattern $Pattern -AllMatches
		write-host $FileContents.Length
		$Update = $false

		ForEach ($Item in $MyMatches)
			{
			$Update = $true
			$ThisMatch = $Item -Match $Pattern
			
			$ReplaceWith = $NewStart + $Matches[1] + $NewEnd
			$FileContents = $FileContents -Replace [regex]::escape($Item), $ReplaceWith
			write-host $Item  "---->"  $ReplaceWith
			}
			If ($Update) {		
				$FileContents|set-content $file
			}
	}
	
<#
If ($file -eq ""){
	Write-host " No File was specified. Please check syntax and try again." `n "Syntax:  .\ReplaceIt.ps1 -file Path:\FileName -find [String] -replace [String] "
break
}

If ($Find -eq "") {
	Write-host '"Find" was not specified. Please check syntax and try again:' `n "Syntax: .\ReplaceIt.ps1 -file Path:\FileName -find [String] -replace [String]"
break
}
 
If ($Replace -eq "") {
	Write-host '"Replace" was not specified. Please check syntax and try again:' `n "Syntax: .\ReplaceIt.ps1 -file Path:\FileName -find [String] -replace [String]"
break
}
#>
	
$FileContents = get-content $file -Erroraction silentlycontinue -ErrorVariable NotFound
If ($NotFound.Count -eq "1"){
	write-host "$file not found. Please check location and filename."
break
}

# $timestamp = get-date -format yyyyMMdd-HHmmss
$SearchString = "$Find"

if (!($AllMatches)) {
$FileContents = ([string]::join("`r`n",$filecontents))

$Found = $fileContents -match "$searchstring"
$logfile = "c:\temp\ReplaceItLog-" + $timestamp + ".txt"	
$newfile = $file + ".bak"

#Try {

	If ($Found){
		$replacewith = $FileContents -replace $SearchString,$Replace
		if (!(test-path $newfile))
		{
			copy-item $file $newfile
		}
		$replacewith|set-content $file
		write-host "Replaced all instances of $find to $Replace in $file and saved backup to $newfile"
		"Replaced all instances of $find to $Replace in $file and saved backup to $newfile" | out-file $logfile -append
		} Else {
			write-host "1>$find not found in $file"
			"$find not found in $file" | out-file $logfile -append
		}

	} 
<#
	catch { 
		"Syntax: ReplaceIt.ps1 -file Path:\FileName -find [String] -replace [String]"
	}

} 
#>
else {
	write-host "Super/Sub"
	SuperScript -FileContents $FileContents -Pattern $Pattern -Start $Start -End $End -NewStart $NewStart -NewEnd $NewEnd -File $File
}