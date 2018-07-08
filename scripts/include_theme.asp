<%


DIM g_sTheme

g_sTheme = "themes/default/"


SUB buildThemeList( aFolders() )

	DIM	sTheme
	DIM	oThemeRoot
	DIM oFolder
	DIM	sName
	DIM	i
	sTheme = findFolder( "themes" )
	SET oThemeRoot = g_oFSO.GetFolder( sTheme )

	i = 0	
	FOR EACH oFolder IN oThemeRoot.SubFolders
		sName = LCASE(oFolder.Name)
		IF "_" <> LEFT(sName,1)  AND  "images" <> sName THEN
			IF UBOUND(aFolders) <= i THEN
				REDIM PRESERVE aFolders(i+5)
			END IF
			sName = virtualFromPhysicalPath( oFolder.Path )
			IF "/" <> RIGHT(sName,1) THEN sName = sName & "/"
			aFolders(i) = sName
			i = i + 1
		END IF
	NEXT 'oFolder
	REDIM PRESERVE aFolders(i-1)
	
	SET oThemeRoot = Nothing
	SET oFolder = Nothing

END SUB



SUB setActiveTheme()
	DIM sFolder
	DIM	sStartFolder
	DIM	sCookieName
	DIM	sSessionKey
	DIM	nCookie
	DIM sCookie

	sCookieName = g_sCookiePrefix & "_Theme_Index"
	sSessionKey = g_sCookiePrefix & "_Theme_Folder"
	
	'sFolder = Session(sSessionKey)
	sFolder = Request.Cookies(sSessionKey)
	sStartFolder = sFolder
	IF "" <> sFolder THEN
		IF NOT g_oFSO.FolderExists(Server.MapPath( sFolder )) THEN
			sFolder = ""
		END IF
	END IF
	IF "" = sFolder THEN
		DIM	j
		DIM	nHigh
		DIM	nLow
		DIM	aFolders()
		REDIM aFolders(5)
		buildThemeList aFolders
		nLow = LBOUND(aFolders)
		nHigh = UBOUND(aFolders)

		sCookie = Request.Cookies( sCookieName )
		IF "" <> sCookie  AND  ISNUMERIC(sCookie) THEN
			nCookie = CINT(sCookie)
			nCookie = nCookie + 1
			IF nHigh < nCookie THEN
				nCookie = nLow
			ELSEIF nCookie < nLow THEN
				nCookie = nLow
			END IF
		ELSE
			RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
			nCookie = ROUND( RND * ( nHigh - nLow )) + nLow
			IF nHigh < nCookie THEN
				nCookie = nHigh
			ELSEIF nCookie < nLow THEN
				nCookie = nLow
			END IF
		END IF
		Response.Cookies( sCookieName ) = nCookie
		Response.Cookies( sCookieName ).Expires = DateAdd( "d", 56, NOW )
		
		sFolder = aFolders(nCookie)

	END IF
	g_sTheme = sFolder
	Session(sSessionKey) = sFolder
	Response.Cookies( sSessionKey ) = sFolder
	
	'Response.Write "start = " & Server.HTMLEncode(sStartFolder) & "<br>" & vbCRLF
	'Response.Write "folder = " & Server.HTMLEncode(sFolder) & "<br>" & vbCRLF
	
END SUB
'setActiveTheme

g_sTheme = "themes/gray/"



%>