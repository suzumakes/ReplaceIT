###Regex

_Backslashes used here instead of backticks, swap them if you need to._

######End all lines with closing tags

    "\r\n" => " \r\n"
    "> \r\n<" => ">\r\n<"
    " \r\n" => " "

Class attributes

    \s+class=[^ >]*

Alignment

    \s+align=[^ >]*

Vertical Alignment

    \s+valign=[^ >]*

Width

    \s+width=[^ >]*

Style

    \s+style='+[^']*'

Span

    </?span+\s+[^>]*>|</span>

Borders

    \s+border=[^ >]*

Cellpadding and spacing

    \s+cellpadding=[^ >]*|\s+cellspacing=[^ >]*

######Combined

    \s+class=[^ >]*|\s+align=[^ >]*|\s+width=[^ >]*|\s+valign=[^ >]*|\s+style='+[^']*'|</?span+\s+[^>]*>|</span>|&nbsp;|<p></p>|\s+border=[^ >]*|\s+cellpadding=[^ >]*|\s+cellspacing=[^ >]*

Pt 2

    <style>(.*`r`n)*</style>|<p></p>|</b><b>|</i><i>|<div>|</div>|<br clear=all>

######Misc

Tag and contents - spans lines

    <tag>(.*`r`n)*</tag>

Everything up to "<"

    [^<]*
