<%



FUNCTION useGalleryFileName( sFileName )

	DIM	sSuffix
	DIM	i
	
	useGalleryFileName = FALSE
	i = InStrRev( sFileName, ".", -1, vbTextCompare )
	IF 0 < i THEN
		sSuffix = LCASE(MID( sFileName, i ))
		SELECT CASE sSuffix
		CASE ".gif", ".jpg", ".jpeg", ".png", ".wmv", ".mpg", ".mov"
			IF "_" = LEFT( sFileName, 1 ) THEN
				useGalleryFileName = FALSE
			ELSE
				useGalleryFileName = TRUE
			END IF
		END SELECT
	END IF

END FUNCTION



DIM aGalleryList
DIM nGalleryCount
'REDIM aGalleryList(5)
nGalleryCount = 0


FUNCTION fetchGalleryFileDescription( oFile )

	fetchGalleryFileDescription = ""
	
	DIM	sPath
	DIM	i
	sPath = oFile.Path
	i = INSTRREV( sPath, "." )
	IF 0 < i THEN
		sPath = LEFT( sPath, i ) & "txt"
		IF g_oFSO.FileExists( sPath ) THEN
			DIM	oDF
			SET oDF = g_oFSO.OpenTextFile( sPath, 1 )
			IF NOT oDF IS Nothing THEN
				DIM	s
				s = oDF.ReadLine
				oDF.Close()
				SET oDF = Nothing
				s = TRIM(s)
				IF "" <> s THEN
					s = REPLACE( s, vbTAB, " " )
					fetchGalleryFileDescription = s
				END IF
			END IF
		END IF
	END IF

END FUNCTION


SUB readGalleryFolderFiles( oFolder, aGallery() )
	DIM oSubFolder
	DIM oFile
	DIM sName
	DIM sTitle
	DIM	sDescription
	DIM sTargetWindow
	DIM	sDate
	DIM bIsIndex
	DIM	sSortName
	DIM sPath
	
	DIM	nGallerySave
	
	nGallerySave = nGalleryCount
	
	sTitle = ""
	sDescription = ""
	sSortName = ""
	sTargetWindow = ""
	sDate = ""

	FOR EACH oFile IN oFolder.Files
		sName = LCase( oFile.Name )
		IF useGalleryFileName( sName ) THEN
			sPath = virtualFromPhysicalPath(oFile.Path)
			sSortName = ""
			IF UBOUND(aGallery) <= nGalleryCount THEN
				REDIM PRESERVE aGallery(nGalleryCount+20)
			END IF
			IF "" <> sSortName THEN
				sSortName = LCASE(sSortName) & sName
			ELSE
				sSortName = sPath
			END IF
			sDescription = fetchGalleryFileDescription( oFile )
			sDate = oFile.DateLastModified
			aGallery(nGalleryCount) = sSortName _
					& vbTAB & sName _
					& vbTAB & sPath _
					& vbTAB & sTitle _
					& vbTAB & sDescription _
					& vbTAB & sDate
			nGalleryCount = nGalleryCount + 1
		END IF
	NEXT 'oFile
	
	FOR EACH oSubFolder IN oFolder.SubFolders
		sName = oSubFolder.Name
		IF "_" <> LEFT(sName,1) THEN
			readGalleryFolderFiles oSubFolder, aGallery
		END IF
	NEXT 'oSubFolder
	
	'IF nGallerySave < nGalleryCount THEN
	'	locationSort aGallery, nGallerySave, nGalleryCount-1
	'END IF
END SUB




SUB readGalleryFolder( sFolder, aGallery() )
	DIM oFolder
	DIM oFile
	DIM sName
	DIM sTitle
	DIM	sDescription
	DIM sTargetWindow
	DIM bIsIndex
	DIM	sSortName
	
	nGalleryCount = 0
	
	sTitle = ""
	sDescription = ""
	sSortName = ""
	sTargetWindow = ""

	IF g_oFSO.FolderExists( sFolder) THEN

		SET oFolder = g_oFSO.GetFolder( sFolder )
	
		readGalleryFolderFiles oFolder, aGallery
	
		IF 0 < nGalleryCount THEN
			REDIM PRESERVE aGallery(nGalleryCount-1)
			shuffle aGallery, 0, nGalleryCount-1
			shuffle aGallery, 0, nGalleryCount-1
			'locationSortDescend aGallery, 0, nGalleryCount-1
		END IF
	END IF

END SUB


SUB buildGalleryIndex( sFolder, sIndexName )

	DIM	aList()
	REDIM aList(50)
	DIM	sStream
	
	readGalleryFolder sFolder, aList
	
	sStream = JOIN( aList, vbCRLF )

	IF g_oFSO.FileExists( sIndexName ) THEN
		g_oFSO.DeleteFile sIndexName, TRUE
	END IF
	
	DIM	oFile
	SET oFile = g_oFSO.CreateTextFile( sIndexName, TRUE )
	IF NOT Nothing IS oFile THEN
		oFile.Write sStream
		oFile.Close
		SET oFile = Nothing
	END IF

END SUB


FUNCTION useGalleryIndex( sIndexName )
	IF g_oFSO.FileExists( sIndexName ) THEN
		DIM	oIndexFile
		SET oIndexFile = g_oFSO.GetFile( sIndexName )
		IF ABS(DATEDIFF( "d", oIndexFile.DateLastModified, NOW )) < 14 THEN
			useGalleryIndex = TRUE
		ELSE
			useGalleryIndex = FALSE
		END IF
	ELSE
		useGalleryIndex = FALSE
	END IF
END FUNCTION




SUB getAllGallery( sFolder, aJokes )

	DIM	sIndexName
	sIndexName = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
	
	IF NOT useGalleryIndex( sIndexName ) THEN
		buildGalleryIndex sFolder, sIndexName
	END IF
	
	DIM	oIndexFile
	SET oIndexFile = g_oFSO.OpenTextFile( sIndexName, 1 )
	
	DIM	sIndexData
	sIndexData = oIndexFile.ReadAll
	oIndexFile.Close
	SET oIndexFile = Nothing
	
	aJokes = SPLIT( sIndexData, vbCRLF )

END SUB




SUB buildGalleryList( sPath )


	getAllGallery sPath, aGalleryList
	
	nGalleryCount = UBOUND(aGalleryList) + 1


END SUB




FUNCTION getGalleryFolder()

	getGalleryFolder = getAppDataFolder( "gallery" )

END FUNCTION




%>