<%
'
'	These functions depend upon the FileSystemObject being created 
'	globally as g_oFSO
'




FUNCTION findFolder( sFolderName )
	DIM	sTemp
	DIM	oFolder
	
	'DIM	g_oFSO
	'SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

	findFolder = ""
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))
	DO UNTIL oFolder.IsRootFolder
		sTemp = g_oFSO.BuildPath( oFolder.Path, sFolderName )
		IF g_oFSO.FolderExists( sTemp ) THEN
			findFolder = sTemp
			EXIT DO
		END IF
		ON Error RESUME Next
		SET oFolder = oFolder.ParentFolder
		IF Err THEN EXIT DO
	LOOP
	ON Error GOTO 0
	SET oFolder = Nothing
	'SET g_oFSO = Nothing

END FUNCTION


FUNCTION findFileUpTree( sFilename )
	DIM	sTemp
	DIM	oFolder
	
	'DIM	g_oFSO
	'SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

	findFileUpTree = ""
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))
	DO UNTIL oFolder.IsRootFolder
		sTemp = g_oFSO.BuildPath( oFolder.Path, sFilename )
		IF g_oFSO.FileExists( sTemp ) THEN
			findFileUpTree = sTemp
			EXIT DO
		END IF
		ON Error RESUME Next
		SET oFolder = oFolder.ParentFolder
		IF Err THEN EXIT DO
	LOOP
	ON Error GOTO 0
	SET oFolder = Nothing
	'SET g_oFSO = Nothing
END FUNCTION


FUNCTION findCGIBIN()
	findCGIBIN = findFolder( "cgi-bin" )
END FUNCTION



FUNCTION findAppData()
	DIM	sFolder
	sFolder = findFolder( "app_data" )
	IF 0 < LEN(sFolder) THEN
		findAppData = sFolder
		EXIT FUNCTION
	END IF
	sFolder = findFolder( "data" )
	IF 0 < LEN(sFolder) THEN
		findAppData = sFolder
		EXIT FUNCTION
	END IF
	sFolder = findFolder( "cgi-bin" )
	IF 0 < LEN(sFolder) THEN
		findAppData = sFolder
		EXIT FUNCTION
	END IF
	findAppData = ""
END FUNCTION




FUNCTION findAppDataFolder( sName )
	DIM	sFolder
	DIM	sPath
	
    sFolder = Server.MapPath("/app_data") & "/" & sName
    IF g_oFSO.FolderExists(sFolder) THEN
        findAppDataFolder = sFolder
        EXIT FUNCTION
    END IF
    sFolder = Server.MapPath("/data") & "/" & sName
    IF g_oFSO.FolderExists(sFolder) THEN
        findAppDataFolder = sFolder
        EXIT FUNCTION
    END IF
    sFolder = Server.MapPath("/cgi-bin") & "/" & sName
    IF g_oFSO.FolderExists(sFolder) THEN
        findAppDataFolder = sFolder
        EXIT FUNCTION
    END IF
    sFolder = Server.MapPath("/" & sName)
    IF g_oFSO.FolderExists(sFolder) THEN
        findAppDataFolder = sFolder
        EXIT FUNCTION
    END IF
	sFolder = findFolder( "app_data\" & sName )
	IF 0 < LEN(sFolder) THEN
		findAppDataFolder = sFolder
		EXIT FUNCTION
	END IF
	sFolder = findFolder( "data\" & sName )
	IF 0 < LEN(sFolder) THEN
		findAppDataFolder = sFolder
		EXIT FUNCTION
	END IF
	sFolder = findFolder( "cgi-bin\" & sName )
	IF 0 < LEN(sFolder) THEN
		findAppDataFolder = sFolder
		EXIT FUNCTION
	END IF
	sFolder = findFolder( sName )
	IF 0 < LEN(sFolder) THEN
		findAppDataFolder = sFolder
		EXIT FUNCTION
	END IF
	findAppDataFolder = ""
END FUNCTION




FUNCTION findfiles_dbFolderSessionKey()
	DIM	i
	DIM	sSessionKey
	sSessionKey = "MDBFolder" & Request.ServerVariables("PATH_INFO")
	i = INSTRREV(sSessionKey,"/")
	IF 0 < i THEN sSessionKey = LEFT(sSessionKey,i-1)
	findfiles_dbFolderSessionKey = sSessionKey
END FUNCTION





FUNCTION findDatabaseFolder()
	DIM	sFolder

	DIM	sSessionKey
	sSessionKey = findfiles_dbFolderSessionKey()

	sFolder = Session.Contents(sSessionKey)
	IF "" <> sFolder THEN
		findDatabaseFolder = sFolder
		EXIT FUNCTION
	END IF

	sFolder = findAppDataFolder( "database" )
	IF 0 < LEN(sFolder) THEN
		Session(sSessionKey) = sFolder
		findDatabaseFolder = sFolder
		EXIT FUNCTION
	END IF
	
	sFolder = findFolder( "database" )
	IF 0 < LEN(sFolder) THEN
		Session(sSessionKey) = sFolder
		findDatabaseFolder = sFolder
		EXIT FUNCTION
	END IF
	findDatabaseFolder = ""
	
END FUNCTION




FUNCTION findProductsFolder()
	DIM	sFolder
	DIM	sPath
	
	sFolder = findAppDataFolder( "products" )
	IF 0 < LEN(sFolder) THEN
		findProductsFolder = sFolder
		EXIT FUNCTION
	END IF
	sFolder = findFolder( "products" )
	IF 0 < LEN(sFolder) THEN
		findProductsFolder = sFolder
		EXIT FUNCTION
	END IF
	findProductsFolder = ""
END FUNCTION


FUNCTION findDBFile( sName )
	DIM	sFolder
	DIM	sPath
	
	DIM	g_oFSO
	SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

	sFolder = findDatabaseFolder()
	IF 0 < LEN( sFolder ) THEN
		sPath = g_oFSO.BuildPath( sFolder, sName )
		IF g_oFSO.FileExists( sPath ) THEN
			findDBFile = sPath
			EXIT FUNCTION
		END IF
	END IF
	sFolder = findProductsFolder()
	IF 0 < LEN( sFolder ) THEN
		sPath = g_oFSO.BuildPath( sFolder, sName )
		IF g_oFSO.FileExists( sPath ) THEN
			findDBFile = sPath
			EXIT FUNCTION
		END IF
	END IF
	findDBFile = ""
	SET g_oFSO = Nothing
END FUNCTION



FUNCTION findXMLFile( sName )
	DIM	sFolder
	DIM	sPath
	
	sFolder = findProductsFolder()
	IF 0 < LEN( sFolder ) THEN
		sPath = g_oFSO.BuildPath( sFolder, sName )
		IF g_oFSO.FileExists( sPath ) THEN
			findXMLFile = sPath
			EXIT FUNCTION
		END IF
	END IF
	findXMLFile = ""
END FUNCTION


FUNCTION findCgibinFile( sFileName )
	DIM	sFolder
	DIM	sPath
	
	sFolder = findFolder( "cgi-bin" )
	IF 0 < LEN( sFolder ) THEN
		sPath = g_oFSO.BuildPath( sFolder, sFileName )
		IF g_oFSO.FileExists( sPath ) THEN
			findCgibinFile = sPath
			EXIT FUNCTION
		END IF
	END IF
	findCgibinFile = ""
END FUNCTION



FUNCTION findCategoryFile()
	findCategoryFile = findXMLFile( "categories.xml" )
END FUNCTION



FUNCTION findProductFile()
	findProductFile = findXMLFile( "products.xml" )
END FUNCTION


FUNCTION createFolderHierarchy( sCacheFolder )
	createFolderHierarchy = TRUE
	IF NOT g_oFSO.FolderExists( sCacheFolder ) THEN
		DIM aTempSplit
		DIM sTemp
		DIM i
		aTempSplit = SPLIT( sCacheFolder, "\", -1, vbTextCompare )
		sTemp = aTempSplit(0)
		FOR i = 1 TO UBOUND(aTempSplit)
			sTemp = sTemp & "\" & aTempSplit(i)
			IF NOT g_oFSO.FolderExists( sTemp ) THEN
				ON ERROR Resume Next
				g_oFSO.CreateFolder( sTemp )
				IF Err THEN createFolderHierarchy = FALSE
				ON ERROR GOTO 0
			END IF
		NEXT 'i
	END IF
END FUNCTION


DIM	g_sRootPath
DIM	g_nLenRootPath
g_sRootPath = Server.MapPath("/")
g_nLenRootPath = LEN(g_sRootPath)


FUNCTION virtualFromPhysicalPath( s )
	virtualFromPhysicalPath = ""
	IF "" <> s THEN
		DIM	sTemp
		sTemp = MID(s,g_nLenRootPath+1)
		IF LCASE(LEFT(s,g_nLenRootPath)) = LCASE(g_sRootPath) THEN
			virtualFromPhysicalPath = REPLACE(sTemp,"\","/",1,-1,vbTextCompare)
		ELSE
			virtualFromPhysicalPath = s
		END IF
	END IF
END FUNCTION





%>