<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2009 Bear Consulting Group
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

FUNCTION htmlFormatForumMeta_f_code( aArgs, sText )
	IF UBOUND(aArgs) < 1 THEN
		htmlFormatForumMeta_f_code = "<code class=""htmlformatforum"">" & sText & "</code>"
	ELSE
		htmlFormatForumMeta_f_code = ""
	END IF
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
		CASE "code"
			htmlFormatMetaSubsFunc = htmlFormatForumMeta_f_code( aFormatSplit, sGenText )
		END SELECT
	END IF

	htmlFormatForum_metaSubsFunc = htmlFormatMetaSubsFunc

END FUNCTION


DIM	g_htmlFormatForum_colors
g_htmlFormatForum_colors = Array( "red", "blue", "pink", "brown", "black", "orange", "violet", _
								"yellow", "green", "gold", "white", "purple" )


FUNCTION htmlFormatForum_tag( sText, sForumTag, sSTFtag )
	DIM	s
	s = REPLACE( sText, "[" & sForumTag & "]", "{{" & sSTFtag & ":" )
	htmlFormatForum_tag = REPLACE( s, "[/" & sForumTag & "]", "}}" )
END FUNCTION

FUNCTION htmlFormatForum_stringReplace( sText )

	DIM	s
	s = sText
	IF 0 < INSTR(s, "[") THEN
		s = htmlFormatForum_tag( s, "b", "b" )
		s = htmlFormatForum_tag( s, "i", "i" )
		s = htmlFormatForum_tag( s, "u", "u" )
		s = htmlFormatForum_tag( s, "s", "s" )
		s = htmlFormatForum_tag( s, "sup", "sup" )
		s = htmlFormatForum_tag( s, "sub", "sub" )
		
		IF 0 < INSTR(s,"[") THEN
			s = htmlFormatForum_tag( s, "big", "big" )
			s = htmlFormatForum_tag( s, "large", "large" )
			s = htmlFormatForum_tag( s, "larger", "larger" )
			s = htmlFormatForum_tag( s, "small", "small" )
			s = htmlFormatForum_tag( s, "smaller", "smaller" )
			DIM	i
			FOR i = LBOUND(g_aFontSizes) TO UBOUND(g_aFontSizes)
				s = htmlFormatForum_tag( s, "size=" & i, "size " & g_aFontSizes(i) )
			NEXT 'i
			s = REPLACE(s,"[/size]","}}")
			
			s = REPLACE(s,"[img]","{{picture ")
			s = REPLACE(s,"[/img]","}}")
			
			
			DIM	sColor
			FOR EACH sColor IN g_htmlFormatForum_colors
				s = htmlFormatForum_tag( s, sColor, sColor )
			NEXT 'sColor
			IF 0 < INSTR(s, "[color=") THEN
				FOR EACH sColor IN g_htmlFormatForum_colors
					s = REPLACE( s, "[color=" & sColor & "]", "{{color " & sColor & ":" )
				NEXT
				s = REPLACE( s, "[/color]", "}}" )
			END IF
		END IF
	END IF
	
	htmlFormatForum_stringReplace = s
	
END FUNCTION



FUNCTION htmlFormatForum_paraReplace( sText )

	DIM	s
	s = htmlFormatForum_stringReplace( sText )
	s = htmlFormatForum_tag( s, "pre", "pre" )
	s = htmlFormatForum_tag( s, "quote", "quote" )
	s = htmlFormatForum_tag( s, "unquote", "unquote" )
	s = htmlFormatForum_tag( s, "code", "code" )
	s = htmlFormatForum_tag( s, "indent", "indent" )
	
	IF 0 < INSTR(s,"[list") THEN
		s = htmlFormatForum_tag( s, "list", "list" )
		s = htmlFormatForum_tag( s, "list=1", "list type=1" )
		s = htmlFormatForum_tag( s, "list=a", "list type=a" )
		s = htmlFormatForum_tag( s, "list=A", "list type=A" )
		s = htmlFormatForum_tag( s, "list=i", "list type=i" )
		s = htmlFormatForum_tag( s, "list=I", "list type=I" )
		s = REPLACE(s,"[/list]","}}")
		s = htmlFormatForum_tag( s, "*", "li" )
		s = htmlFormatForum_tag( s, "li", "li" )
	END IF
	IF 0 < INSTR(s,"[table") THEN
		s = htmlFormatForum_tag( s, "table", "table border=1" )
		s = htmlFormatForum_tag( s, "tr", "tr" )
		s = htmlFormatForum_tag( s, "td", "td" )
		s = htmlFormatForum_tag( s, "th", "th" )
	END IF

	s = htmlFormatForum_tag( s, "left", "align left" )
	s = htmlFormatForum_tag( s, "center", "align center" )
	s = htmlFormatForum_tag( s, "right", "align right" )
	
	s = REPLACE(s,"[hr]", "{{hr}}")
	s = REPLACE(s,"[br]", "{{br}}")
	
	DIM	n
	FOR n = 1 TO 6
		s = htmlFormatForum_tag( s, "h" & n, "header " & n )
	NEXT
	htmlFormatForum_paraReplace = s
	
END FUNCTION


FUNCTION htmlFormatForumString( sText )

	DIM	s
	s = htmlFormatForum_stringReplace( sText )

	'g_htmlFormatForum_metaSubsFunc = g_htmlFormat_metaSubsFunc
	'g_htmlFormat_metaSubsFunc = "htmlFormatForum_metaSubsFunc"

	htmlFormatForumString = htmlFormatMeta( s )

	'g_htmlFormat_metaSubsFunc = g_htmlFormatForum_metaSubsFunc
	
END FUNCTION


FUNCTION htmlFormatForum( sText )
	DIM	sString
	sString = htmlFormatForum_paraReplace( sText )
	
	'g_htmlFormatForum_metaSubsFunc = g_htmlFormat_metaSubsFunc
	'g_htmlFormat_metaSubsFunc = "htmlFormatForum_metaSubsFunc"

	htmlFormatForum = htmlFormatCRLF( sString )
	
	'g_htmlFormat_metaSubsFunc = g_htmlFormatForum_metaSubsFunc

END FUNCTION


%>