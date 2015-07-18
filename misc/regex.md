###Regex

Remove any tag and its contents

	<tag>(.*`r`n)*</tag>

Bring tags to the same line

	> `r`n<"

Remove class attributes

	\s+class=[^ >]*

Remove alignment

	\s+align=[^ >]*

Remove width

	\s+width=[^ >]*

Remove vertical alignment

	\s+valign=[^ >]*

Remove styles

	\s+style='+[^']*'

Remove span tags

	</?span+\s+[^>]*>

Remove borders

	\s+border=[^ >]*

Remove cellpadding for tables

	\s+cellpadding=[^ >]*

Remove cellspacing for tables

	\s+cellspacing=[^ >]*

