<%


FUNCTION xmldom_make()
	DIM	oXML
	SET oXML = Server.CreateObject("Msxml2.DOMDocument.6.0")
	IF NOT oXML IS Nothing THEN
		oXML.async = false
		oXML.setProperty "ProhibitDTD", false
		oXML.setProperty "ResolveExternals", true
		oXML.setProperty "AllowDocumentFunction", true
		'oXML.setProperty "AllowXmlAttributes", true
		oXML.setProperty "ValidateOnParse", true
	ELSE
		Response.Write "Problem creating XML object<br>"
	END IF
	SET xmldom_make = oXML
	SET oXML = Nothing
END FUNCTION





FUNCTION xmldom_loadFile( sFile )

	SET xmldom_loadFile = Nothing

	DIM	oXML
	SET oXML = xmldom_make()
	IF 0 < INSTR(LCASE(sFile),".xslt") THEN
		oXML.setProperty "AllowXsltScript", true
	END IF
	oXML.load( sFile )

	IF oXML.parseError.errorCode <> 0 THEN
		Response.Write "Error in File: " & Server.HTMLEncode(sFile) & "<br>" & vbCRLF
		Response.Write "Error Code: " & oXML.parseError.errorCode & "<br>" & vbCRLF
		Response.Write "Error Reason: " & oXML.parseError.reason & "<br>" & vbCRLF
		Response.Write "Error Line: " & oXML.parseError.line & "<br>" & vbCRLF
	ELSE
		SET xmldom_loadFile = oXML
	END IF
	
	SET oXML = Nothing

END FUNCTION


FUNCTION xmldom_loadXML( sString )

	SET xmldom_loadXML = Nothing

	DIM	oXML
	SET oXML = xmldom_make()
	oXML.loadXML( sString )

	IF oXML.parseError.errorCode <> 0 THEN
		Response.Write "Error in File: " & Server.HTMLEncode(sFile) & "<br>" & vbCRLF
		Response.Write "Error Code: " & oXML.parseError.errorCode & "<br>" & vbCRLF
		Response.Write "Error Reason: " & oXML.parseError.reason & "<br>" & vbCRLF
		Response.Write "Error Line: " & oXML.parseError.line & "<br>" & vbCRLF
	ELSE
		SET xmldom_loadXML = oXML
	END IF
	
	SET oXML = Nothing

END FUNCTION


FUNCTION xmldom_makeHTTP()
	SET xmldom_makeHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
END FUNCTION


FUNCTION xmldom_fetchXMLHTTP( sURL )
	SET xmldom_fetchXMLHTTP = Nothing
	DIM	oHTTP
	SET oHTTP = xmldom_makeHTTP()
	IF NOT oHTTP IS Nothing THEN
		oHTTP.Open "GET", sURL, False
		oHTTP.SetRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")
		oHTTP.Send
		IF 0 = Err.Number THEN
			DIM	nStatus
			nStatus = CINT(oHTTP.Status)
			IF 200 = nStatus  OR  0 = nStatus THEN
			
				'Response.Write "nStatus = " & nStatus & "<br>"
				'Response.Flush
			
				DIM	sNewMime
				
				sNewMime = oHTTP.getResponseHeader( "Content-Type" )
				'Response.Write "Mime = " & sNewMime & "<br>"
				SELECT CASE LCASE(sNewMime)
				CASE "text/xml"
					SET xmldom_fetchXMLHTTP = oHTTP.ResponseXML
				END SELECT
							
			END IF
		END IF
	END IF
	SET oHTTP = Nothing
END FUNCTION






%>