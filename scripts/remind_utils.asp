<%
' dependencies
'
'	findfiles.asp
'



DIM	g_remind_sDataFolder
g_remind_sDataFolder = ""


FUNCTION findRemindFolder()
	DIM	sFolder
	
	IF "" = g_remind_sDataFolder THEN
		sFolder = findAppDataFolder( "remind" )
		IF "" <> sFolder THEN
			findRemindFolder = sFolder
			g_remind_sDataFolder = sFolder
		ELSE
			findRemindFolder = ""
		END IF
	ELSE
		findRemindFolder = g_remind_sDataFolder
	END IF
END FUNCTION


FUNCTION getRemindSortname( sName )
	DIM	i
	i = INSTR(sName,";sort=")
	IF 0 < i THEN
		DIM	sTemp
		sTemp = MID(sName,i+6)
		i = INSTR(sTemp,";")
		IF 0 < i THEN
			sTemp = LEFT(sTemp,i-1)
		ELSE
			sTemp = LEFT(sTemp,LEN(sTemp)-4)
		END IF
		i = INSTR(sName,";")
		IF 0 < i THEN
			getRemindSortname = sTemp & LEFT(sName,i-1)
		ELSE
			getRemindSortname = sTemp & LEFT(sName,LEN(sName)-4)
		END IF
	ELSE
		getRemindSortname = sName
	END IF
END FUNCTION


SUB getRemindList( aList() )
	DIM	sFolder
	DIM	oFolder
	sFolder = findRemindFolder()
	IF "" <> sFolder THEN
		SET oFolder = g_oFSO.GetFolder( sFolder )
		IF NOT Nothing IS oFolder THEN
			DIM	nCount
			DIM	oFile
			DIM	sName
			DIM	sNameLC
			DIM	sFile
			DIM	sFileLC
			nCount = 0
			FOR EACH oFile IN oFolder.Files
				sName = oFile.Name
				sNameLC = LCASE(sName)
				IF ".xml" = RIGHT(sNameLC,4) THEN
					sFile = oFile.Path
					IF UBOUND(aList) <= nCount THEN
						REDIM PRESERVE aList(nCount+5)
					END IF
					aList(nCount) = getRemindSortname(sNameLC) & vbTAB & sFile & vbTAB & CSTR(oFile.DateLastModified)
					nCount = nCount + 1
				END IF
			NEXT 'sFile
			REDIM PRESERVE aList(nCount-1)
			sort aList, LBOUND(aList), UBOUND(aList)
		END IF
	END IF
END SUB


FUNCTION findRemindFile( sRemindName )
	DIM	sFolder
	DIM	sPath
	
	sFolder = findRemindFolder()
	IF 0 < LEN( sFolder ) THEN
		sPath = g_oFSO.BuildPath( sFolder, sRemindName )
		IF g_oFSO.FileExists( sPath ) THEN
			findRemindFile = sPath
			EXIT FUNCTION
		END IF
	END IF
	findRemindFile = ""
END FUNCTION


FUNCTION remindCSS()

	remindCSS = "scripts/remind_css.asp"

END FUNCTION



FUNCTION remindMakeXmlDOM()
	SET remindMakeXMLDOM = xmldom_make()
END FUNCTION



FUNCTION remindLoadXmlFile( sFile )

	SET remindLoadXmlFile = xmldom_loadFile( sFile )

END FUNCTION


FUNCTION remindLoadWithIncludeHref( sHREF )
	SET remindLoadWithIncludeHref = Nothing
	IF "~" = LEFT(sHREF,1) THEN
		DIM	i
		DIM	sURL
		DIM	sFilename
		DIM	sQuery
		sHREF = MID(sHREF,2)
		IF "/" = LEFT(sHREF,1) THEN sHREF = MID(sHREF,2)
		i = INSTR(sHREF,"?")
		IF 0 < i THEN
			sFilename = LEFT(sHREF,i-1)
			sQuery = MID(sHREF,i)
		ELSE
			sFilename = sHREF
			sQuery = ""
		END IF
		sURL = findFileUpTree( sFilename )
		IF "" <> sURL THEN
			sHREF = "http://" & Request.ServerVariables("HTTP_HOST") & virtualFromPhysicalPath( sURL ) & sQuery
		END IF
	END IF

	IF "" <> sHREF THEN
		DIM oHTTP
		SET oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
		IF NOT oHTTP IS Nothing THEN
			'Response.Write sHREF & "<br>"
			'Response.Flush
			oHTTP.Open "GET", sHREF, False
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
						SET remindLoadWithIncludeHref = oHTTP.ResponseXML
					END SELECT
								
				END IF
			END IF

			SET oHTTP = Nothing
		END IF
	END IF
END FUNCTION


FUNCTION attValue( v )
	IF ISNULL( v ) THEN
		attValue = ""
	ELSEIF ISEMPTY( v ) THEN
		attValue = ""
	ELSE
		attValue = v
	END IF
END FUNCTION



FUNCTION remindLoadWithIncludeCache( sFile, sInterval, nValue, sBreak )
	SET remindLoadWithIncludeCache = Nothing
	
	IF "" = attValue(sFile) THEN EXIT FUNCTION
	IF "" = attValue(sInterval) THEN EXIT FUNCTION
	IF "" = attValue(nValue) THEN EXIT FUNCTION
	IF ISNUMERIC(nValue) THEN
		IF CLNG(nValue) < 1 THEN EXIT FUNCTION
	ELSE
		EXIT FUNCTION
	END IF
	
	DIM	oFile
	SET oFile = cache_openTextFile( "remind", sFile, sInterval, nValue, sBreak )
	IF NOT oFile IS Nothing THEN
		DIM	sXMLData
		sXMLData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
		DIM oXML
		SET oXML = xmldom_loadXML( sXMLData )
		IF NOT oXML IS Nothing THEN
			SET remindLoadWithIncludeCache = oXML
			
			SET oXML = Nothing
		END IF
	END IF

END FUNCTION



FUNCTION remindLoadWithInclude( sFile )

	SET remindLoadWithInclude = Nothing
	
	DIM	oXML
	SET oXML = remindLoadXmlFile( sFile )
	IF NOT oXML IS Nothing THEN
	
		DIM	oInc
		SET oInc = oXML.selectSingleNode("/reminders/event")
		IF oInc IS Nothing THEN
		
			SET oInc = oXML.selectSingleNode("/reminders/include")
			IF NOT oInc IS Nothing THEN
		
				DIM	sHREF
				sHREF = oInc.getAttribute("href")
				IF "" <> sHREF THEN
					DIM	sCacheFile
					DIM	sCacheInterval
					DIM	sCacheValue
					DIM	sCacheBreak
					sCacheFile = oInc.getAttribute("cache-file")
					sCacheInterval = oInc.GetAttribute("cache-interval")
					sCacheValue = oInc.GetAttribute("cache-value")
					sCacheBreak = oInc.GetAttribute("cache-break-interval")
					DIM sXMLData
					DIM	otXML
					DIM	oFile
					SET otXML = remindLoadWithIncludeCache( sCacheFile, sCacheInterval, sCacheValue, sCacheBreak)
					IF otXML IS Nothing THEN
						SET otXML = remindLoadWithIncludeHref( sHREF )
						IF NOT otXML IS Nothing THEN
							sXMLData = otXML.xml
							SET oFile = cache_makeFile( "remind", sCacheFile )
							IF NOT oFile IS Nothing THEN
								oFile.Write sXMLData
								oFile.Close
								SET oFile = Nothing
							END IF
						END IF
					END IF
					SET remindLoadWithInclude = otXML
					SET otXML = Nothing
				END IF

			END IF
		
			SET oInc = Nothing
		
		ELSE
		
			SET remindLoadWithInclude = oXML
		
		END IF
		
		SET oXML = Nothing
	
	END IF

END FUNCTION



CLASS CRemindFileServer

	DIM	m_sFolder
	DIM	m_oFolder

	SUB Class_Initialize()
		m_sFolder = ""
		SET m_oFolder = Nothing
	END SUB

	SUB Class_Terminate()
		SET m_oFolder = Nothing
	END SUB
	
	PUBLIC PROPERTY GET Count()
	END PROPERTY
	
	PUBLIC PROPERTY GET Collection( ByVal sKey )
	END PROPERTY

END CLASS


%>