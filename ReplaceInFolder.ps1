<#

This is a Find and Replace PowerShell Script for cleaning Word-Generated HTML.

Many thanks to Michael Clark for the inital replaceit PowerShell Script.

====================
TO DO

1. Table Styling
2. Test Lists
3. Find more accurate Super and SubScript Tagging

====================

#>

Param (
	$Folder
)
Function DoIt
	{
		Param (
			$Folder
		)
		$Folders = Get-ChildItem $Folder
		ForEach ($Child in $Folders)
			{
				If ($Child.PSIsContainer)
					{
						DoIt $Child.FullName
					}
				Else
					{
						# Arrange all lines to end with closing tags
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n" -Replace " `r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "> `r`n<" -Replace ">`r`n<"
						.\replaceit.ps1 -File $Child.FullName -Find " `r`n" -Replace " "
						.\replaceit.ps1 -File $Child.FullName -Find "</span>" -Replace "</span>`r`n"
						
						# Standardize SuperScript Tags
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)2.5pt'>" -Replace "position:relative;top:-4.5pt'>"
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)4.0pt'>" -Replace "position:relative;top:-4.5pt'>"
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)4.5pt'>" -Replace "position:relative;top:-4.5pt'>"
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)5.0pt'>" -Replace "position:relative;top:-4.5pt'>"

						# Insert SuperScripts
						$Start = "position:relative;top:-4.5pt'>"
						$End = "</span>"
						$Pattern = $Start + "(.*?)" + $End
						$NewStart = "position:relative;top:-4.5pt'>" + "<sup>"
						$NewEnd = "</sup>" + "</span>"
						.\replaceit.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd

						# Remove leftover "position:relative;top:-4.5pt'>"
						# .\replaceit.ps1 -File $Child.FullName -Find "position:relative;top:-4.5pt'>" -Replace ""
<#
						# Original SuperScript Tag Search
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)pt;(.*)letter-spacing:(.*)pt'>" -Replace "position:relative;top:-4.5pt'>"
#>
						# Standardize SubScript Tags
						# position:relative;top[^>]*> <-- Dangerous! This finds everything leftover!
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)2.0pt'>" -Replace "position:relative;top:2.0pt'>"
						# .\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)3.5t'>" -Replace "position:relative;top:2.0pt'>"
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)3.0pt'>" -Replace "position:relative;top:2.0pt'>"
						.\replaceit.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)5.5pt'>" -Replace "position:relative;top:2.0pt'>"

						# Insert SubScripts
						$Start = "position:relative;top:2.0pt'>"
						$End = "</span>"
						$Pattern = $Start + "(.*?)" + $End
						$NewStart = "position:relative;top:2.0pt'>" + "<sub>"
						$NewEnd = "</sub>" + "</span>"
						.\replaceit.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd

						# Remove leftover "position:relative;top:2.0pt'>" 
						# .\replaceit.ps1 -File $Child.FullName -Find "position:relative;top:2.0pt'>" -Replace ""
<#
						# (Original) Insert SubScripts
						$Start = "'font-size:8.0pt;font-family:" + '"Times New Roman"' + "," +'"serif"' + "'>"
						$End = "</span>"
						$Pattern = $Start + "(.*?)" + $End
						$NewStart = "'font-size:8.0pt;font-family:" + '"Times New Roman"' + "," +'"serif"' + "'>" + "<sub>"
						$NewEnd = "</sub>" + "</span>"
						.\replaceit.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd
#>
						.\replaceit.ps1 -File $Child.FullName -Find "</span>`r`n" -Replace "</span>"

						# Removes class, align, width, and style attributes, borders and cellpadding and spacing, span tags, and empty <p> tags
						.\replaceit.ps1 -File $Child.FullName -Find "\s+class=[^ >]*|\s+align=[^ >]*|\s+width=[^ >]*|\s+valign=[^ >]*|\s+style='+[^']*'|</?span+\s+[^>]*>|</span>|&nbsp;|<p></p>|\s+border=[^ >]*|\s+cellpadding=[^ >]*|\s+cellspacing=[^ >]*" -Replace ""

						# Removes Style Declaration
						.\replaceit.ps1 -File $Child.FullName -Find "<style>(.*`r`n)*</style>" -Replace ""

						# Removes Leftover Empty <p> tags, and nested <b> and <i> tags
						.\replaceit.ps1 -File $Child.FullName -Find "<p></p>|</b><b>|</i><i>" -Replace ""

						# Removes Divs and Breaks
						.\replaceit.ps1 -File $Child.FullName -Find "<div>|</div>|<br clear=all>" -Replace ""

						# Change <b> and <i> to <strong> and <em>
						.\replaceit.ps1 -File $Child.FullName -Find "<i>" -Replace "<em>"
						.\replaceit.ps1 -File $Child.FullName -Find "</i>" -Replace "</em>"
						.\replaceit.ps1 -File $Child.FullName -Find "<b>" -Replace "<strong>"
						.\replaceit.ps1 -File $Child.FullName -Find "</b>" -Replace "</strong>"
						
						# M$ Spacing Character - Change -Replace to " " to maintain Spacing
						.\replaceit.ps1 -File $Child.FullName -Find " " -Replace ""

						# ASCII codes to HTML
						.\replaceit.ps1 -File $Child.FullName -Find "®" -Replace "&#174;"
						.\replaceit.ps1 -File $Child.FullName -Find "©" -Replace "&#169;"
						.\replaceit.ps1 -File $Child.FullName -Find "µ" -Replace "&#181;"
						
						# M$ Specific ASCII to HTML
						.\replaceit.ps1 -File $Child.FullName -Find "€" -Replace "&#8364;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#128;" -Replace "&#8364;"

						# .\replaceit.ps1 -File $Child.FullName -Find "&#129;" -Replace "THIS CODE NOT USED"

						.\replaceit.ps1 -File $Child.FullName -Find "‚" -Replace "&#8218;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#130;" -Replace "&#8218;"
						.\replaceit.ps1 -File $Child.FullName -Find "ƒ" -Replace "&#402;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#131;" -Replace "&#402;"
						.\replaceit.ps1 -File $Child.FullName -Find '„' -Replace "&#8222;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#132;" -Replace "&#8222;"
						.\replaceit.ps1 -File $Child.FullName -Find "…" -Replace "&#8230;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#133;" -Replace "&#8230;"
						.\replaceit.ps1 -File $Child.FullName -Find "†" -Replace "&#8224;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#134;" -Replace "&#8224;"
						.\replaceit.ps1 -File $Child.FullName -Find "‡" -Replace "&#8225;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#135;" -Replace "&#8225;"
						.\replaceit.ps1 -File $Child.FullName -Find "ˆ" -Replace "&#710;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#136;" -Replace "&#710;"
						.\replaceit.ps1 -File $Child.FullName -Find "‰" -Replace "&#8240;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#137;" -Replace "&#8240;"
						.\replaceit.ps1 -File $Child.FullName -Find "Š" -Replace "&#352;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#138;" -Replace "&#352;"
						.\replaceit.ps1 -File $Child.FullName -Find "‹" -Replace "&#8249;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#139;" -Replace "&#8249;"
						.\replaceit.ps1 -File $Child.FullName -Find "Œ" -Replace "&#338;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#140;" -Replace "&#338;"

						# .\replaceit.ps1 -File $Child.FullName -Find "&#141;" -Replace "THIS CODE NOT USED"

						.\replaceit.ps1 -File $Child.FullName -Find "Ž" -Replace "&#381;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#142;" -Replace "&#381;"

						# .\replaceit.ps1 -File $Child.FullName -Find "&#143;" -Replace "THIS CODE NOT USED"
						# .\replaceit.ps1 -File $Child.FullName -Find "&#144;" -Replace "THIS CODE NOT USED"

						# Comment out the following line to retain curly Word single quotes
						# .\replaceit.ps1 -File $Child.FullName -Find "‘|’" -Replace "'"
						.\replaceit.ps1 -File $Child.FullName -Find "‘" -Replace "&#8216;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#145;" -Replace "&#8216;"
						.\replaceit.ps1 -File $Child.FullName -Find "’" -Replace "&#8217;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#146;" -Replace "&#8217;"

						# Comment out the following line to retain curly Word double quoutes
						# .\replaceit.ps1 -File $Child.FullName -Find '“|”' -Replace '"'
						.\replaceit.ps1 -File $Child.FullName -Find '“' -Replace "&#8220;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#147;" -Replace "&#8220;"
						.\replaceit.ps1 -File $Child.FullName -Find '”' -Replace "&#8221;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#148;" -Replace "&#8221;"

						.\replaceit.ps1 -File $Child.FullName -Find "•" -Replace "&#8226;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#149;" -Replace "&#8226;"
<#
						# Bullets to Lists - In Progress
						.\replaceit.ps1 -File $Child.FullName -Find "<p>•" -Replace "</li><li>"
						.\replaceit.ps1 -File $Child.FullName -Find "•" -Replace "</li><li>"
#>
<#
						# Replace "<p>&#8226;" with Lists
						$Start = "<p>&#8226;"
						$End = "</p>"
						$Pattern = $Start + "(.*?)" + $End
						$NewStart = "<li>"
						$NewEnd = "</li>"
						.\replaceit.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd
#>
						.\replaceit.ps1 -File $Child.FullName -Find "–" -Replace "&#8211;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#150;" -Replace "&#8211;"
						.\replaceit.ps1 -File $Child.FullName -Find "—" -Replace "&#8212;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#151;" -Replace "&#8212;"
						.\replaceit.ps1 -File $Child.FullName -Find "˜" -Replace "&#732;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#152;" -Replace "&#732;"
						.\replaceit.ps1 -File $Child.FullName -Find "™" -Replace "&#8482;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#153;" -Replace "&#8482;"
						.\replaceit.ps1 -File $Child.FullName -Find "š" -Replace "&#353;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#154;" -Replace "&#353;"
						.\replaceit.ps1 -File $Child.FullName -Find "›" -Replace "&#8250;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#155;" -Replace "&#8250;"
						.\replaceit.ps1 -File $Child.FullName -Find "œ" -Replace "&#339;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#156;" -Replace "&#339;"

						# .\replaceit.ps1 -File $Child.FullName -Find "&#157;" -Replace "THIS CODE NOT USED"

						.\replaceit.ps1 -File $Child.FullName -Find "ž" -Replace "&#382;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#158;" -Replace "&#382;"
						.\replaceit.ps1 -File $Child.FullName -Find "Ÿ" -Replace "&#376;"
						.\replaceit.ps1 -File $Child.FullName -Find "&#159;" -Replace "&#376;"

						# Remove Stubborn Line Breaks
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n `r`n" -Replace "`r`n"

						# Remove Excessive Line Breaks
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
						.\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"

						# Uncomment to Remove more Empty Lines
						# .\replaceit.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
						
						# Remove leftover "position:relative;top:-4.5pt'>"
						.\replaceit.ps1 -File $Child.FullName -Find "position:relative;top:-4.5pt'>" -Replace ""
						# Remove leftover "position:relative;top:2.0pt'>" 
						.\replaceit.ps1 -File $Child.FullName -Find "position:relative;top:2.0pt'>" -Replace ""

					}
			}
	}

DoIt $Folder
