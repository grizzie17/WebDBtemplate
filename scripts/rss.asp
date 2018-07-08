<%

SUB rss_fetchHTTP( sContent, sMime, sURL, sQuery, bBinary )

	DIM	nStatus
	DIM	oHTTP
	
	sContent = ""
	bBinary = FALSE
	sMime = ""

	ON ERROR Resume Next
	SET oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	IF Nothing IS oHTTP THEN
		ON ERROR GOTO 0
		sContent = ""
		EXIT SUB
	END IF
	
	'oHTTP.Open "HEAD", sURL, False
	'oHTTP.Send
	'nStatus = Err.Number
	'IF 0 = nStatus THEN
	
	'	nStatus = CINT(oHTTP.Status)
		
		'IF 200 = nStatus  OR  202 = nStatus  OR  500 = nStatus THEN
	
			oHTTP.Open "GET", sURL, False
			'oHTTP.SetRequestHeader "Referer", "http://localhost/test.asp"
			oHTTP.SetRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")
			oHTTP.Send
			IF 0 = Err.Number THEN
				nStatus = CINT(oHTTP.Status)
				
				IF 200 = nStatus  OR  0 = nStatus THEN
				
					DIM	sNewMime
					
					sNewMime = oHTTP.getResponseHeader( "Content-Type" )
					'Response.Write "Mime = " & sNewMime & "<br>"
					bBinary = FALSE
					SELECT CASE sNewMime
					CASE ""
						sContent = oHTTP.ResponseXML.xml
						IF LEN(sContent) < 2 THEN
							sContent = oHTTP.ResponseText
						END IF
						sMime = "text/xml"
					CASE "text/xml"
						sContent = oHTTP.ResponseXML.xml
						sMime = "text/xml"
					CASE "text/txt"
						sContent = oHTTP.ResponseText
						sMime = "text/txt"
					CASE "image/gif"
						bBinary = TRUE
						sContent = oHTTP.responseBody
						sMime = "image/gif"
					CASE ELSE
						sContent = oHTTP.ResponseText
						sMime = "text/txt"
					END SELECT
								
				END IF
			END IF
		
		'END IF
	'ELSE
	'	Response.Write "rss_fetchHTTP Err = " & HEX(nStatus) & "<br>"
	'	nStatus = CINT(oHTTP.Status)
	'	Response.Write "http status = " & HEX(nStatus) & "<br>"
	'END IF
	
	SET oHTTP = Nothing
	ON Error GOTO 0

END SUB




FUNCTION rss_getXML( sLocal, sURL, sRefer )
	SET rss_getXML = rss_getXMLCached( sLocal, sURL, sRefer, "h", 2, "d" )
END FUNCTION





FUNCTION rss_getXMLCached( sLocal, sURL, sRefer, sInterval, sIntValue, sBreakInterval )

	SET rss_getXMLCached = Nothing

	DIM	sFolder
	
	DIM	bFetch
	bFetch = FALSE
	
	DIM	oFile
	DIM sXMLFile
	sXMLFile = cache_checkFile( "rss", sLocal, sInterval, sIntValue, sBreakInterval )
	IF "" = sXMLFile THEN
		sXMLFile = cache_filepath( "rss", sLocal )
		bFetch = TRUE
	ELSE
		bFetch = FALSE
	END IF
	
	DIM oXMLFile
	
	
	IF bFetch THEN
	
		DIM	sContent
		DIM	sMime
		DIM	bBinary
		
		rss_fetchHTTP sContent, sMime, sURL, "", bBinary
		
		IF 0 < LEN(sContent) THEN
		
			IF LEN(sContent) < 20 THEN
				sXMLFile = ""
			ELSE
			
				IF g_oFSO.FileExists( sXMLFile ) THEN
					g_oFSO.DeleteFile sXMLFile, TRUE
				END IF
				
			END IF
		
		END IF

	ELSE
	

		IF g_oFSO.FileExists( sXMLFile ) THEN
		
			SET oFile = g_oFSO.OpenTextFile( sXMLFile, 1 )
			IF NOT Nothing IS oFile THEN
				sContent = oFile.ReadAll
				oFile.Close
				SET oFile = Nothing
			END IF
		
		END IF

	END IF
	
	IF "" <> sXMLFile  AND  0 < LEN(sContent) THEN
	
		DIM	oXML
		SET oXML = Server.CreateObject("msxml2.DOMDocument.6.0")
		
		oXML.async = false
		oXML.resolveExternals = false
		oXML.setProperty "ProhibitDTD", false
		oXML.setProperty "ResolveExternals", true
		oXML.setProperty "AllowDocumentFunction", true
		oXML.setProperty "AllowXsltScript", true

		oXML.loadXML( sContent )
		IF oXML.parseError.errorCode <> 0  AND  -1072898035 <> oXML.parseError.errorCode THEN
			Response.Write "Error in XML file: " & sLocal & "<br>" & vbCRLF
			Response.Write "- Code: " & oXML.parseError.errorCode & "<br>" & vbCRLF
			Response.Write "- Reason: " & oXML.parseError.reason & "<br>" & vbCRLF
			Response.Write "- Line: " & oXML.parseError.line & "<br>" & vbCRLF
			SET oXML = Nothing
			EXIT FUNCTION
		ELSE
			'Response.Write "<p>bFetch = " & bFetch & "; sXMLFile = " & Server.HTMLEncode(sXMLFile) & "</p>"
			IF bFetch  AND  "" <> sXMLFile THEN
				ON ERROR RESUME NEXT
				oXML.save sXMLFile
				'Response.Write "<p>" & Err.number & "</p>"
				ON ERROR GOTO 0
			END IF
		END IF
		SET rss_getXMLCached = oXML
		SET oXML = Nothing
	END IF

END FUNCTION


%>