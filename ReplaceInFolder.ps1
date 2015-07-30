<#

This is a Find and Replace PowerShell Script for cleaning Word-Filtered HTML.

Many thanks to Michael Clark.

==============================
TO DO

1. Format bulleted lists
2. Even better super/subscript checking?
==============================

#>

Param (
  $Folder,
  [switch]$Log
)

Set-Variable -Name LogIT -value $Log -scope Global

$Folders = Get-ChildItem $Folder

ForEach ( $Child in $Folders ) {
  write-host "$Child"
}

$Response = read-host "convert these files? (Y/n)"

If ( $Response -eq "" -or $Response -eq "y" -or $Response -eq "Y" ) {

  Function DoIt {

    Param ( $Folder )

    $Folders = Get-ChildItem $Folder

    ForEach ( $Child in $Folders ) {

      If ( $Child.PSIsContainer ) {
        DoIt $Child.FullName
      }

      Else {

        $Extensions = "/.htm/"

        $FileExtension = "/" + $Child.Extension + "/"

        If ( $Extensions.Contains( $FileExtension ) -and $FileExtension -gt "" ) {

          # arrange all lines to end with closing tags
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n" -Replace " `r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "> `r`n<" -Replace ">`r`n<"
          .\ReplaceIT.ps1 -File $Child.FullName -Find " `r`n" -Replace " "
          .\ReplaceIT.ps1 -File $Child.FullName -Find "<span" -Replace "`r`n<span"

          # standardize superscript tags
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)2.5(.*)'>" -Replace "insertsuper'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)4.0(.*)'>" -Replace "insertsuper'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)4.5(.*)'>" -Replace "insertsuper'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)5.0(.*)'>" -Replace "insertsuper'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)5.5(.*)'>" -Replace "insertsuper'>"

          # insert superscripts
          $Start = "insertsuper'>"
          $End = "</span>"
          $Pattern = $Start + "(.*?)" + $End
          $NewStart = "insertsuper'><sup>"
          $NewEnd = "</sup></span>"
          .\ReplaceIT.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd

          # standardize subscript tags
          # .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)1.5(.*)'>" -Replace "insertsub'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)2.0(.*)'>" -Replace "insertsub'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "position:relative;top:(.*)3.0(.*)'>" -Replace "insertsub'>"

          # insert subscripts
          $Start = "insertsub'>"
          $End = "</span>"
          $Pattern = $Start + "(.*?)" + $End
          $NewStart = "insertsub'><sub>"
          $NewEnd = "</sub></span>"
          .\ReplaceIT.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd

          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n<span" -Replace "<span"

          # removes class, align, width, and style attributes, borders and cellpadding and spacing, span tags, and empty <p> tags
          .\ReplaceIT.ps1 -File $Child.FullName -Find "\s+class=[^ >]*|\s+align=[^ >]*|\s+width=[^ >]*|\s+valign=[^ >]*|\s+style='+[^']*'|</?span+\s+[^>]*>|</span>|&nbsp;|<p></p>|\s+border=[^ >]*|\s+cellpadding=[^ >]*|\s+cellspacing=[^ >]*" -Replace ""

          # removes style declaration, leftover empty and nested <p>, <b>, and <i> tags, divs, and breaks
          .\ReplaceIT.ps1 -File $Child.FullName -Find "<style>(.*`r`n)*</style>|<p></p>|</b><b>|</i><i>|<div>|</div>|<br clear=all>" -Replace ""

          # change <b> and <i> to <strong> and <em>
          .\ReplaceIT.ps1 -File $Child.FullName -Find "<i>" -Replace "<em>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "</i>" -Replace "</em>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "<b>" -Replace "<strong>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "</b>" -Replace "</strong>"

          # M$ spacing character
          .\ReplaceIT.ps1 -File $Child.FullName -Find " " -Replace ""

          # M$ specific ASCII to HTML 128 - 159
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

          # uncomment to eliminate single "Smart Quotes"
          # .\ReplaceIT.ps1 -File $Child.FullName -Find "‘|’" -Replace "'"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "‘" -Replace "&#8216;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "&#145;" -Replace "&#8216;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "’" -Replace "&#8217;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "&#146;" -Replace "&#8217;"

          # uncomment to eliminate "Smart Quotes"
          # .\ReplaceIT.ps1 -File $Child.FullName -Find '“|”' -Replace '"'
          .\ReplaceIT.ps1 -File $Child.FullName -Find '“' -Replace "&#8220;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "&#147;" -Replace "&#8220;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find '”' -Replace "&#8221;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "&#148;" -Replace "&#8221;"

          # bullets
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
          # END M$ specific ASCII to HTML 128 - 159

          # remove line breaks
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n `r`n" -Replace "`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n`r`n" -Replace "`r`n`r`n"

          # fewer line breaks
          # .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n`r`n" -Replace "`r`n"

          # ASCII codes to HTML
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¡" -Replace "&#161;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¢" -Replace "&#162;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "£" -Replace "&#163;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¤" -Replace "&#164;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¥" -Replace "&#165;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¦" -Replace "&#166;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "§" -Replace "&#167;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¨" -Replace "&#168;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "©" -Replace "&#169;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "ª" -Replace "&#170;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "«" -Replace "&#171;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¬" -Replace "&#172;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "­" -Replace "&#173;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "®" -Replace "&#174;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¯" -Replace "&#175;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "°" -Replace "&#176;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "±" -Replace "&#177;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "²" -Replace "&#178;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "³" -Replace "&#179;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "´" -Replace "&#180;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "µ" -Replace "&#181;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¶" -Replace "&#182;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "·" -Replace "&#183;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¸" -Replace "&#184;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¹" -Replace "&#185;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "º" -Replace "&#186;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "»" -Replace "&#187;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¼" -Replace "&#188;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "½" -Replace "&#189;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¾" -Replace "&#190;"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "¿" -Replace "&#191;"

          # foreign language characters.
          .\CMatch.ps1 -File $Child.FullName -Find "À" -Replace "&#192;"
          .\CMatch.ps1 -File $Child.FullName -Find "Á" -Replace "&#193;"
          .\CMatch.ps1 -File $Child.FullName -Find "Â" -Replace "&#194;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ã" -Replace "&#195;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ä" -Replace "&#196;"
          .\CMatch.ps1 -File $Child.FullName -Find "Å" -Replace "&#197;"
          .\CMatch.ps1 -File $Child.FullName -Find "Æ" -Replace "&#198;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ç" -Replace "&#199;"
          .\CMatch.ps1 -File $Child.FullName -Find "È" -Replace "&#200;"
          .\CMatch.ps1 -File $Child.FullName -Find "É" -Replace "&#201;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ê" -Replace "&#202;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ë" -Replace "&#203;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ì" -Replace "&#204;"
          .\CMatch.ps1 -File $Child.FullName -Find "Í" -Replace "&#205;"
          .\CMatch.ps1 -File $Child.FullName -Find "Î" -Replace "&#206;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ï" -Replace "&#207;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ð" -Replace "&#208;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ñ" -Replace "&#209;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ò" -Replace "&#210;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ó" -Replace "&#211;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ô" -Replace "&#212;"
          .\CMatch.ps1 -File $Child.FullName -Find "Õ" -Replace "&#213;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ö" -Replace "&#214;"
          .\CMatch.ps1 -File $Child.FullName -Find "×" -Replace "&#215;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ø" -Replace "&#216;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ù" -Replace "&#217;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ú" -Replace "&#218;"
          .\CMatch.ps1 -File $Child.FullName -Find "Û" -Replace "&#219;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ü" -Replace "&#220;"
          .\CMatch.ps1 -File $Child.FullName -Find "Ý" -Replace "&#221;"
          .\CMatch.ps1 -File $Child.FullName -Find "Þ" -Replace "&#222;"
          .\CMatch.ps1 -File $Child.FullName -Find "ß" -Replace "&#223;"
          .\CMatch.ps1 -File $Child.FullName -Find "à" -Replace "&#224;"
          .\CMatch.ps1 -File $Child.FullName -Find "á" -Replace "&#225;"
          .\CMatch.ps1 -File $Child.FullName -Find "â" -Replace "&#226;"
          .\CMatch.ps1 -File $Child.FullName -Find "ã" -Replace "&#227;"
          .\CMatch.ps1 -File $Child.FullName -Find "ä" -Replace "&#228;"
          .\CMatch.ps1 -File $Child.FullName -Find "å" -Replace "&#229;"
          .\CMatch.ps1 -File $Child.FullName -Find "æ" -Replace "&#230;"
          .\CMatch.ps1 -File $Child.FullName -Find "ç" -Replace "&#231;"
          .\CMatch.ps1 -File $Child.FullName -Find "è" -Replace "&#232;"
          .\CMatch.ps1 -File $Child.FullName -Find "é" -Replace "&#233;"
          .\CMatch.ps1 -File $Child.FullName -Find "ê" -Replace "&#234;"
          .\CMatch.ps1 -File $Child.FullName -Find "ë" -Replace "&#235;"
          .\CMatch.ps1 -File $Child.FullName -Find "ì" -Replace "&#236;"
          .\CMatch.ps1 -File $Child.FullName -Find "í" -Replace "&#237;"
          .\CMatch.ps1 -File $Child.FullName -Find "î" -Replace "&#238;"
          .\CMatch.ps1 -File $Child.FullName -Find "ï" -Replace "&#239;"
          .\CMatch.ps1 -File $Child.FullName -Find "ð" -Replace "&#240;"
          .\CMatch.ps1 -File $Child.FullName -Find "ñ" -Replace "&#241;"
          .\CMatch.ps1 -File $Child.FullName -Find "ò" -Replace "&#242;"
          .\CMatch.ps1 -File $Child.FullName -Find "ó" -Replace "&#243;"
          .\CMatch.ps1 -File $Child.FullName -Find "ô" -Replace "&#244;"
          .\CMatch.ps1 -File $Child.FullName -Find "õ" -Replace "&#245;"
          .\CMatch.ps1 -File $Child.FullName -Find "ö" -Replace "&#246;"
          .\CMatch.ps1 -File $Child.FullName -Find "÷" -Replace "&#247;"
          .\CMatch.ps1 -File $Child.FullName -Find "ø" -Replace "&#248;"
          .\CMatch.ps1 -File $Child.FullName -Find "ù" -Replace "&#249;"
          .\CMatch.ps1 -File $Child.FullName -Find "ú" -Replace "&#250;"
          .\CMatch.ps1 -File $Child.FullName -Find "û" -Replace "&#251;"
          .\CMatch.ps1 -File $Child.FullName -Find "ü" -Replace "&#252;"
          .\CMatch.ps1 -File $Child.FullName -Find "ý" -Replace "&#253;"
          .\CMatch.ps1 -File $Child.FullName -Find "þ" -Replace "&#254;"
          .\CMatch.ps1 -File $Child.FullName -Find "ÿ" -Replace "&#255;"

          # remove leftover "insertsuper'>|insertsub'>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "insertsuper'>|insertsub'>" -Replace ""

          # replace M$ images with placeholder
          .\ReplaceIT.ps1 -File $Child.FullName -Find '<img(.*)">' -Replace '<img class="myimgclass" src="images/image000.jpg" alt="" title="">'
          .\ReplaceIT.ps1 -File $Child.FullName -Find '<p><img(.*)"></p>' -Replace '<img class="myimgclass" src="images/image000.jpg" alt="" title="">'

          # basic table formatting
          .\ReplaceIT.ps1 -File $Child.FullName -Find '<table>' -Replace '<table border="1" align="center" cellpadding="3" cellspacing="0">'
          .\ReplaceIT.ps1 -File $Child.FullName -Find "</p>\s\s\s<p>" -Replace " "
          .\ReplaceIT.ps1 -File $Child.FullName -Find "<td>\s+<p>" -Replace "`r`n<td>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "</p>\s+</td>" -Replace "</td>"

          # combine super/subscript tags
          .\ReplaceIT.ps1 -File $Child.FullName -Find "</sup>`r`n<sup>" -Replace ""
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n<sup>" -Replace "<sup>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find "</sub>`r`n<sub>" -Replace ""
          .\ReplaceIT.ps1 -File $Child.FullName -Find "`r`n<sub>" -Replace "<sub>"
          .\ReplaceIT.ps1 -File $Child.FullName -Find '</sup>\)' -Replace ')</sup>'
          .\ReplaceIT.ps1 -File $Child.FullName -Find '\(<sup>' -Replace '<sup>('
          .\ReplaceIT.ps1 -File $Child.FullName -Find '</sub>\)' -Replace ')</sub>'
          .\ReplaceIT.ps1 -File $Child.FullName -Find '\(<sub>' -Replace '</sub>('
          
          # bullets to lists
          .\ReplaceIT.ps1 -File $Child.FullName -Find '<p>&#183;' -Replace '<li>'
          .\ReplaceIT.ps1 -File $Child.FullName -Find '<p>&#8226;' -Replace '<li>'

          $Start = "<li>"
          $End = "</p>"
          $Pattern = $Start + "(.*?)" + $End
          $NewStart = "<li>"
          $NewEnd = "</li>"
          .\replaceit.ps1 -File $Child.FullName -AllMatches -Start $Start -End $End -Pattern $Pattern -NewStart $NewStart -NewEnd $NewEnd
        }
      }
    }
  }
} Else {

  write-host "nothing converted"
  break

}

DoIt $Folder

