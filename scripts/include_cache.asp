<%




DIM	g_sCacheBaseFolder
g_sCacheBaseFolder = ""


FUNCTION cache_getBaseFolder()
	cache_getBaseFolder = g_sCacheBaseFolder
	IF "" = cache_getBaseFolder THEN
		g_sCacheBaseFolder = findAppDataFolder( "cache" )
		IF "" = g_sCacheBaseFolder THEN
			g_sCacheBaseFolder = findAppData()
			g_sCacheBaseFolder = g_sCacheBaseFolder & "\cache"
			createFolderHierarchy g_sCacheBaseFolder
		END IF
		cache_getBaseFolder = g_sCacheBaseFolder
	END IF
END FUNCTION

FUNCTION cache_getFolder( sFolder )
	DIM	sBaseFolder
	DIM	sPath
	sBaseFolder = cache_getBaseFolder()
	sPath = g_oFSO.BuildPath( sBaseFolder, sFolder )
	IF g_oFSO.FolderExists( sPath ) THEN
		cache_getFolder = sPath
	ELSE
		createFolderHierarchy sPath
		cache_getFolder = sPath
	END IF
END FUNCTION



FUNCTION cache_filepath( sCache, sFile )

	DIM	sFolder
	sFolder = cache_getFolder( sCache )
	
	DIM	sPath
	sPath = g_oFSO.BuildPath( sFolder, sFile )
	
	cache_filepath = sPath
	
END FUNCTION



'
' returns filespecification if we can use the cache; returns empty string otherwise
'
FUNCTION cache_checkFile( sCache, sFile, sInterval, nIntValue, sBreakInterval )

	cache_checkFile = ""	' assume that we can't get the file

	DIM	sPath
	sPath = cache_filepath( sCache, sFile )
	
	IF g_oFSO.FileExists( sPath ) THEN
		DIM	oFile
		SET oFile = g_oFSO.GetFile( sPath )
		IF 20 < oFile.Size THEN		' we need the file to be larger than 20 bytes
			DIM	bForce
			DIM	dLastMod
			bForce = FALSE
			dLastMod = oFile.DateLastModified
			SELECT CASE LCASE(sBreakInterval)
			CASE "d"	' must refresh on day boundry
				IF DATEVALUE(NOW) <> DATEVALUE(dLastMod) THEN bForce = TRUE
			CASE "w", "ww"
				IF DATEPART("ww",NOW)&"-"&YEAR(NOW) <> DATEPART("ww",dLastMod)&"-"&DATEPART("ww",dLastMod) THEN bForce = TRUE
			CASE "m"	' must refresh on month boundry
				IF MONTH(NOW)&"-"&YEAR(NOW) <> MONTH(dLastMod)&"-"&YEAR(dLastMod) THEN bForce = TRUE
			CASE "y", "yyyy"
				IF YEAR(NOW) <> YEAR(dLastMod) THEN bForce = TRUE
			END SELECT
			IF NOT bForce THEN
				IF ABS(DATEDIFF( sInterval, dLastMod, NOW )) <= nIntValue THEN
					cache_checkFile = sPath
				END IF
			END IF
		END IF
		SET oFile = Nothing
	END IF

END FUNCTION



FUNCTION cache_openTextFile( sCache, sFile, sInterval, nIntValue, sBreakInterval )

	SET cache_openTextFile = Nothing	' assume that we can't get the file

	DIM	sPath
	sPath = cache_checkFile( sCache, sFile, sInterval, nIntValue, sBreakInterval )
	
	IF "" <> sPath THEN
		SET cache_openTextFile = g_oFSO.OpenTextFile( sPath, 1, 0 )
	END IF
	
END FUNCTION




' deprecated function
FUNCTION cache_getFile( sCache, sFile, sInterval, nIntValue, sBreakInterval )

	SET cache_getFile = cache_openTextFile( sCache, sFile, sInterval, nIntValue, sBreakInterval )

END FUNCTION



FUNCTION cache_makeFile( sCache, sFile )

	DIM	sPath
	sPath = cache_filepath( sCache, sFile )
	
	IF g_oFSO.FileExists( sPath ) THEN
		g_oFSO.DeleteFile( sPath )
	END IF
	
	SET cache_makeFile = g_oFSO.CreateTextFile( sPath, TRUE )
	
END FUNCTION




SUB cache_clearSubFolders( oFolder )

	DIM	sName
	DIM	oSub
	FOR EACH oSub in oFolder.SubFolders
		sName = oSub.name
		IF "_" <> LEFT(sName,1) THEN
			cache_clearSubFolders oSub
			oSub.Delete( false )
		END IF
	NEXT 'oSub
	
	DIM	oFile
	FOR EACH oFile IN oFolder.Files
		oFile.Delete( true )
	NEXT 'oFile

END SUB




SUB cache_clearFolders()

	DIM	sBase
	sBase = cache_getBaseFolder()
	
	IF g_oFSO.FolderExists( sBase ) THEN
	
		DIM	oBase
		SET oBase = g_oFSO.GetFolder( sBase )
		
		IF NOT oBase IS Nothing THEN
		
			cache_clearSubFolders oBase
		
		END IF
	
	END IF

END SUB


SUB cache_clearFolder( sCache )

	DIM	sFolder
	sFolder = cache_getFolder( sCache )
	
	IF g_oFSO.FolderExists( sFolder ) THEN
	
		DIM	oFolder
		SET oFolder = g_oFSO.GetFolder( sFolder )
		
		IF NOT oFolder IS Nothing THEN
		
			cache_clearSubFolders oFolder
		
		END IF
	
	END IF

END SUB


SUB cache_deleteFile( sCache, sFile )

	DIM	sPath
	sPath = cache_filepath( sCache, sFile )
	
	IF g_oFSO.FileExists( sPath ) THEN
		g_oFSO.DeleteFile( sPath )
	END IF

END SUB


%>