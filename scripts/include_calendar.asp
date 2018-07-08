<%




FUNCTION calendar_buildLink( sProtocol, sURL, sTarget, sText )
	IF "address:" = sProtocol THEN
		DIM	sMapURL
		sMapURL = "map.asp?q=" & sURL
		IF "" = sText  OR  (sProtocol & sURL) = sText THEN
			calendar_buildLink = htmlFormat_f_makeLink( "", sMapURL, sTarget, REPLACE(sURL,"+"," ") )
		ELSE
			calendar_buildLink = htmlFormat_f_makeLink( "", sMapUrl, sTarget, sText )
		END IF
	ELSEIF "local:" = sProtocol THEN
		DIM sLocalURL
		sLocalURL = sURL
		calendar_buildLink = htmlFormat_f_makeLink( "", sLocalURL, "_self", sText )
	ELSE
		calendar_buildLink = htmlFormat_f_makeLink( sProtocol, sURL, sTarget, sText )
	END IF
END FUNCTION

g_htmlFormat_makeLinkFunc = "calendar_buildLink"





%>