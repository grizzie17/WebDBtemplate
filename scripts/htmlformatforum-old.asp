<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2007 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------
'
'	Similar to Simple Text Format these routines format based upon the
'	popular forum codes:
'
'		Text Styles:
'			[b]bold[/b]
'			[i]italic[/i]
'			[u]underline[/u]
'			[s]strike-through[/s]
'
'		Font Colors:
'			[color-name]yada[/color-name]
'			supported colors: 
'				red, blue, pink, brown, black, orange, violet, yellow, green,
'				gold, white, purple
'
'		Font Sizes
'			[size=1]text[/size=1]	x-small
'			[size=2]text[/size=2]	small
'			[size=3]text[/size=3]	medium
'			[size=4]text[/size=4]	large
'			[size=5]text[/size=5]	x-large
'			[size=6]text[/size=6]	xx-large
'			[big]text[/big]
'			[small]text[/small]
'			[smaller]text[/smaller]
'
'		Headings:
'			[h1]text[/h1]
'			[h2]text[/h2]
'			[h3]text[/h3]
'			[h4]text[/h4]
'			[h5]text[/h5]
'			[h6]text[/h6]
'
'		Text alignment:
'			[left]yada[/left]
'			[center]yada[/center]
'			[right]yada[/center]
'
'		Bullet List:
'			[list]
'			[*]text[/*]
'			{*]text[/*]
'			[/list]
'
'		Ordered Alpha List:
'			[list=a]
'			[*]text[/*]
'			[/list=a]
'
'		Ordered Number List:
'			[list=1]
'			[*]text[/*]
'			[/list=1]
'
'		Code:
'			[code]text[/code]
'
'		Pre:
'			[pre]text[/pre]
'
'		Quote:
'			[quote]text[/quote]
'
'		Images:
'			[img]url[/img]
'
'		Links (URLs):
'			www.domainname.com
'			http://domainname.tld
'			[url=some-url]link text[/url]
'

DIM g_htmlFormatForum_metaSubsFunc
g_htmlFormatForum_metaSubsFunc = ""

DIM	g_aFontSizes
g_aFontSizes = ARRAY( "xx-small", "x-small", "small", "medium", "large", "x-large", "xx-large" )


FUNCTION htmlFormatForumMeta_f_pre( aArgs, sText )
	IF UBOUND(aArgs) < 1 THEN
		htmlFormatForumMeta_f_pre = "<pre class=""htmlformatforum"">" & sText & "</pre>"
	ELSE
		htmlFormatForumMeta_f_pre = ""
	END IF
END FUNCTION

FUNCTION htmlFormatForumMeta_f_quote( aArgs, sText )
	IF UBOUND(aArgs) < 1 THEN
		htmlFormatForumMeta_f_quote = "<cite class=""htmlformatforum"">" & sText & "</cite>"
	ELSE
		htmlFormatForumMeta_f_quote = ""
	END IF
END FUNCTION

FUNCTION htmlFormatForumMeta_f_code( aArgs, sText )
	IF UBOUND(aArgs) < 1 THEN
		htmlFormatForumMeta_f_code = "<code class=""htmlformatforum"">" & sText & "</code>"
	ELSE
		htmlFormatForumMeta_f_code = ""
	END IF
END FUNCTION

FUNCTION htmlFormatForumMeta_f_list( aArgs, sText )
END FUNCTION

FUNCTION htmlFormatForumMeta_f_listelement( aArgs, sText )
END FUNCTION

FUNCTION htmlFormatForumMeta_f_left( aArgs, sText )
END FUNCTION

FUNCTION htmlFormatForumMeta_f_center( aArgs, sText )
	IF UBOUND(aArgs) < 1 THEN
		htmlFormatForumMeta_f_center = "<center class=""htmlformatforum"">" & sText & "</center>"
	ELSE
		htmlFormatForumMeta_f_center = ""
	END IF
END FUNCTION

FUNCTION htmlFormatForumMeta_f_right( aArgs, sText )
END FUNCTION


FUNCTION htmlFormatForum_metaSubsFunc( aFormatSplit(), sGenText )

	DIM	htmlFormatMetaSubsFunc

	IF "" <> g_htmlFormatForum_metaSubsFunc THEN
		htmlFormatMetaSubsFunc = EVAL( g_htmlFormatForum_metaSubsFunc & "( aFormatSplit, sGenText )" )
	ELSE
		htmlFormatMetaSubsFunc = ""
	END IF
	IF "" = htmlFormatMetaSubsFunc THEN

		SELECT CASE LCASE( aFormatSplit(0) )
		CASE "pre"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_pre( aFormatSplit, sGenText )
		CASE "quote"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_quote( aFormatSplit, sGenText )
		CASE "code"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_code( aFormatSplit, sGenText )
		CASE "list"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_list( aFormatSplit, sGenText )
		CASE "*"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_listelement( aFormatSplit, sGenText )
		CASE "left"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_left( aFormatSplit, sGenText )
		CASE "center"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_center( aFormatSplit, sGenText )
		CASE "right"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_right( aFormatSplit, sGenText )
		END SELECT
	END IF

	htmlFormatForum_metaSubsFunc = htmlFormatMetaSubsFunc

END FUNCTION


DIM	g_htmlFormatForum_colors
g_htmlFormatForum_colors = Array( "red", "blue", "pink", "brown", "black", "orange", "violet", _
								"yellow", "green", "gold", "white", "purple" )


FUNCTION htmlFormatForum_stringReplace( sText )

	DIM	s
	s = sText
	s = REPLACE(s,"[b]","{{b:")
	s = REPLACE(s,"[/b]","}}")
	s = REPLACE(s,"[i]","{{i:")
	s = REPLACE(s,"[/i]","}}")
	s = REPLACE(s,"[u]","{{u:")
	s = REPLACE(s,"[/u]","}}")
	s = REPLACE(s,"[s]","{{s:")
	s = REPLACE(s,"[/s]","}}")
	s = REPLACE(s,"[sup]","{{sup:")
	s = REPLACE(s,"[/sup]","}}")
	s = REPLACE(s,"[sub]","{{sub:")
	s = REPLACE(s,"[/sub]","}}")
	
	s = REPLACE(s,"[big]","{{big:")
	s = REPLACE(s,"[/big]","}}")
	s = REPLACE(s,"[large]","{{large:")
	s = REPLACE(s,"[/large]","}}")
	s = REPLACE(s,"[larger]","{{larger:")
	s = REPLACE(s,"[/larger]","}}")
	s = REPLACE(s,"[small]","{{small:")
	s = REPLACE(s,"[/small]","}}")
	s = REPLACE(s,"[smaller]","{{smaller:")
	s = REPLACE(s,"[/smaller]","}}")
	
	s = REPLACE(s,"[img]","{{picture ")
	s = REPLACE(s,"[/img]","}}")
	
	DIM	i
	FOR i = LBOUND(g_aFontSizes) TO UBOUND(g_aFontSizes)
		s = REPLACE(s,"[size="&i&"]","{{size "&g_aFontSizes(i)&":")
		s = REPLACE(s,"[/size="&i&"]","}}")
	NEXT 'i
	
	DIM	sColor
	FOR EACH sColor IN g_htmlFormatForum_colors
		s = REPLACE( s, "["&sColor&"]", "{{"&sColor&":" )
		s = REPLACE( s, "[/"&sColor&"]", "}}" )
	NEXT 'sColor
	
	htmlFormatForum_stringReplace = s
	
END FUNCTION



FUNCTION htmlFormatForum_paraReplace( sText )

	DIM	s
	s = htmlFormatForum_stringReplace( sText )
	s = REPLACE(s,"[pre]","{{pre:")
	s = REPLACE(s,"[/pre]","}}")
	s = REPLACE(s,"[quote]","{{quote:")
	s = REPLACE(s,"[/quote]","}}")
	s = REPLACE(s,"[code]","{{code:")
	s = REPLACE(s,"[/code]","}}")
	s = REPLACE(s,"[list]","{{list:")
	s = REPLACE(s,"[/list]","}}")
	s = REPLACE(s,"[list=1]","{{list type=1:")
	s = REPLACE(s,"[/list=1]","}}")
	s = REPLACE(s,"[list=a]","{{list type=a:")
	s = REPLACE(s,"[/list=a]","}}")
	s = REPLACE(s,"[*]","{{*:")
	s = REPLACE(s,"[/*]","}}")
	s = REPLACE(s,"[left]","{{left:")
	s = REPLACE(s,"[/left]","}}")
	s = REPLACE(s,"[center]","{{center:")
	s = REPLACE(s,"[/center]","}}")
	s = REPLACE(s,"[right]","{{right:")
	s = REPLACE(s,"[/right]","}}")
	htmlFormatForum_paraReplace = s
	
END FUNCTION


FUNCTION htmlFormatForumString( sText )

	DIM	s
	s = htmlFormatForum_stringReplace( sText )

	g_htmlFormatForum_metaSubsFunc = g_htmlFormat_metaSubsFunc
	g_htmlFormat_metaSubsFunc = "htmlFormatForum_metaSubsFunc"

	htmlFormatForumString = htmlFormatMeta( s )

	g_htmlFormat_metaSubsFunc = g_htmlFormatForum_metaSubsFunc
	
END FUNCTION


FUNCTION htmlFormatForum( sText )
	DIM	sString
	sString = htmlFormatForum_paraReplace( sText )
	
	g_htmlFormatForum_metaSubsFunc = g_htmlFormat_metaSubsFunc
	g_htmlFormat_metaSubsFunc = "htmlFormatForum_metaSubsFunc"

	htmlFormatForum = htmlFormatCRLF( sString )
	
	g_htmlFormat_metaSubsFunc = g_htmlFormatForum_metaSubsFunc

END FUNCTION


%>