<#

This is a Find and Replace PowerShell Script for cleaning Word-Generated HTML.

Many Thanks for Michael Clark for being a baller.

====================
TO DO

1. Tables
2. Lists
3. Super and Subscripts - better than current 90%
4. Implement user confirmation - i.e. "would you like to run ReplaceIT on file_a?"

====================

#>

Param ( $Folder )

Function DoIt {

  Param ( $Folder )

  $Folders = Get-ChildItem $Folder

  ForEach ($Child in $Folders) {

    If ($Child.PSIsContainer) {
      DoIt $Child.FullName
    }

    Else {

      $Extensions = "/.txt/.htm/"

      $FileExtension = "/" + $Child.Extension + "/"

      If ($Extensions.Contains($FileExtension) -and $FileExtension -gt "") {

        # Arrange all lines to end with closing tags
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n" -Replace " `r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "> `r`n<" -Replace ">`r`n<"
        .\ReplaceIT.ps1 -File $Child.FullName -Find " `r`n" -Replace " "
        .\ReplaceIT.ps1 -File $Child.FullName -Find "</span>" -Replace "</span>`r`n"

        # Standardize Superscript Tags
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)2.5pt'>" -Replace "position:relative;top:-4.5pt'>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)4.0pt'>" -Replace "position:relative;top:-4.5pt'>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)4.5pt'>" -Replace "position:relative;top:-4.5pt'>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)5.0pt'>" -Replace "position:relative;top:-4.5pt'>"

        # Insert Superscripts
        $Start = "position:relative;top:-4.5pt'>"
        $End = "</span>"
        $Pattern = $Start + "(.*?)" + $End
        $NewStart = "position:relative;top:-4.5pt'>" + "<sup>"
        $NewEnd = "</sup>" + "</span>"
        .\ReplaceIT.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd

        # Standardize Subscript Tags
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)2.0pt'>" -Replace "position:relative;top:2.0pt'>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)3.0pt'>" -Replace "position:relative;top:2.0pt'>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:(.*)relative;(.*)top:(.*)5.5pt'>" -Replace "position:relative;top:2.0pt'>"

        # Insert Subscripts
        $Start = "position:relative;top:2.0pt'>"
        $End = "</span>"
        $Pattern = $Start + "(.*?)" + $End
        $NewStart = "position:relative;top:2.0pt'>" + "<sub>"
        $NewEnd = "</sub>" + "</span>"
        .\ReplaceIT.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd

        .\ReplaceIT.ps1 -File $Child.FullName -Find "</span>`r`n" -Replace "</span>"

        # Removes class, align, width, and style attributes, borders and cellpadding and spacing, span tags, and empty <p> tags
        .\ReplaceIT.ps1 -File $Child.FullName -Find "\s+class=[^ >]*|\s+align=[^ >]*|\s+width=[^ >]*|\s+valign=[^ >]*|\s+style='+[^']*'|</?span+\s+[^>]*>|</span>|&nbsp;|<p></p>|\s+border=[^ >]*|\s+cellpadding=[^ >]*|\s+cellspacing=[^ >]*" -Replace ""

        # Removes Style Declaration
        .\ReplaceIT.ps1 -File $Child.FullName -Find "<style>(.*`r`n)*</style>" -Replace ""

        # Removes Leftover Empty <p> tags, and nested <b> and <i> tags
        .\ReplaceIT.ps1 -File $Child.FullName -Find "<p></p>|</b><b>|</i><i>" -Replace ""

        # Removes Divs and Breaks
        .\ReplaceIT.ps1 -File $Child.FullName -Find "<div>|</div>|<br clear=all>" -Replace ""

        # Change <b> and <i> to <strong> and <em>
        .\ReplaceIT.ps1 -File $Child.FullName -Find "<i>" -Replace "<em>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "</i>" -Replace "</em>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "<b>" -Replace "<strong>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "</b>" -Replace "</strong>"

        # M$ Spacing Character - Change -Replace to " " to maintain Spacing
        .\ReplaceIT.ps1 -File $Child.FullName -Find " " -Replace ""

        # ASCII codes to HTML
        .\ReplaceIT.ps1 -File $Child.FullName -Find "®" -Replace "&#174;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "©" -Replace "&#169;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "µ" -Replace "&#181;"

        # M$ Specific ASCII to HTML
        .\ReplaceIT.ps1 -File $Child.FullName -Find "€" -Replace "&#8364;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#128;" -Replace "&#8364;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "‚" -Replace "&#8218;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#130;" -Replace "&#8218;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "ƒ" -Replace "&#402;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#131;" -Replace "&#402;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find '„' -Replace "&#8222;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#132;" -Replace "&#8222;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "…" -Replace "&#8230;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#133;" -Replace "&#8230;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "†" -Replace "&#8224;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#134;" -Replace "&#8224;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "‡" -Replace "&#8225;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#135;" -Replace "&#8225;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "ˆ" -Replace "&#710;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#136;" -Replace "&#710;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "‰" -Replace "&#8240;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#137;" -Replace "&#8240;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "Š" -Replace "&#352;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#138;" -Replace "&#352;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "‹" -Replace "&#8249;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#139;" -Replace "&#8249;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "Œ" -Replace "&#338;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#140;" -Replace "&#338;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "Ž" -Replace "&#381;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#142;" -Replace "&#381;"

        # Uncomment the line below to eliminate single "Smart Quotes"
        # .\ReplaceIT.ps1 -File $Child.FullName -Find "‘|’" -Replace "'"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "‘" -Replace "&#8216;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#145;" -Replace "&#8216;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "’" -Replace "&#8217;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#146;" -Replace "&#8217;"

        # Uncomment the line below to eliminate "Smart Quotes"
        # .\ReplaceIT.ps1 -File $Child.FullName -Find '“|”' -Replace '"'
        .\ReplaceIT.ps1 -File $Child.FullName -Find '“' -Replace "&#8220;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#147;" -Replace "&#8220;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find '”' -Replace "&#8221;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#148;" -Replace "&#8221;"

        # Bullets
        .\ReplaceIT.ps1 -File $Child.FullName -Find "•" -Replace "&#8226;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#149;" -Replace "&#8226;"

        .\ReplaceIT.ps1 -File $Child.FullName -Find "–" -Replace "&#8211;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#150;" -Replace "&#8211;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "—" -Replace "&#8212;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#151;" -Replace "&#8212;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "˜" -Replace "&#732;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#152;" -Replace "&#732;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "™" -Replace "&#8482;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#153;" -Replace "&#8482;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "š" -Replace "&#353;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#154;" -Replace "&#353;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "›" -Replace "&#8250;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#155;" -Replace "&#8250;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "œ" -Replace "&#339;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#156;" -Replace "&#339;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "ž" -Replace "&#382;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#158;" -Replace "&#382;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "Ÿ" -Replace "&#376;"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "&#159;" -Replace "&#376;"

        # Fix eacute
        .\ReplaceIT.ps1 -File $Child.FullName -Find "é" -Replace "&eacute;"

        # Remove Stubborn Line Breaks
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n `r`n" -Replace "`r`n"

        # Remove Excessive Line Breaks
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"

        # Uncomment to Remove more Line Breaks
        # .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"

        # Remove leftover "position:relative;top:-4.5pt'>"
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:-4.5pt'>" -Replace ""
        # Remove leftover "position:relative;top:2.0pt'>" 
        .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:2.0pt'>" -Replace ""
      } 
    }
  }
}

DoIt $Folder

