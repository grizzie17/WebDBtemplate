<%


FUNCTION forum_formatHtml( sCacheFile, sDomain, sContentName, sTitle, nMinutes )

	forum_formatHtml = ""

	DIM	oXML
	SET oXML = rss_getXMLCached( sCacheFile, _
					"http://forum." & sDomain & "/smf/index.php?action=.xml;type=rss2", "", _
					"n", nMinutes, "d" )
	IF NOT Nothing IS oXML THEN

		DIM oXSL
		SET oXSL = xmldom_loadFile( Server.MapPath( "scripts/rssforum_atom.xslt" ) )
		IF NOT oXSL IS Nothing THEN
			DIM sText
			sText = oXML.transformNode(oXSL)
			IF 0 < LEN(sText) THEN
				sText = REPLACE( sText, "_$$", "&" )
				forum_formatHtml =	"" _
					&	"<div id=""" & sContentName & "content"" class=""rssForum dynamic"">" & vbCRLF _
					&		"<div class=""BlockHead"">" & vbCRLF _
					&			"<a target=""_blank"" href=""http://forum." & sDomain & """>" & sTitle & "</a>" & vbCRLF _
					&		"</div>" & vbCRLF _
					&		"<div class=""BlockBody"">" & vbCRLF _
					&		"<ul>" & vbCRLF _
					&			sText _
					&		"</ul>" & vbCRLF _
					&		"</div>" & vbCRLF _
					&	"</div>" & vbCRLF
			END IF
			SET oXSL = Nothing
		END IF
		
		SET oXML = Nothing

	END IF

END FUNCTION


SUB RSSForum( sCacheFile, sDomain, sIcon, sTitle, nMinutes )

	DIM	sHtml
	sHtml = forum_formatHtml( sCacheFile, sDomain, sIcon, sTitle, nMinutes )
	IF 0 < LEN(sHtml) THEN
		Response.Write sHtml
	END IF
END SUB

%>