<%


DIM	g_sPicturePath
g_sPicturePath = ""

DIM g_Upload
SET g_Upload = Nothing

SUB picture_createObject()
	SET g_Upload = Server.CreateObject("ASPSimpleUpload.Upload")
	IF Nothing IS g_Upload THEN
		Response.Write "Problem creating upload component<br>"
		Response.Flush
		'Response.End
	END IF
END SUB


SUB picture_close()
	SET g_Upload = Nothing
END SUB


FUNCTION Requestor( s )
	picture_debug s
	IF Nothing IS g_Upload THEN
		Requestor = Request( s )
		picture_debug " : "
	ELSE
		Requestor = g_Upload.Form( s )
		picture_debug " = "
	END IF
	picture_debug Requestor & "<br>"
END FUNCTION





SUB picture_debug( s )
	'Response.Write s & vbCRLF
	'Response.Flush
END SUB


SUB picture_findPath( sFolder )
	DIM	sPath
	sPath = findDatabaseFolder()
	sPath = g_oFSO.BuildPath( sPath, sFolder )
	IF createFolderHierarchy( sPath ) THEN
		g_sPicturePath = sPath
	ELSE
		g_sPicturePath = ""
	END IF
END SUB


SUB picture_delete( sFile )

	DIM	sPictPath
	
	sPictPath = g_oFSO.BuildPath( g_sPicturePath, sFile )
	IF g_oFSO.FileExists( sPictPath ) THEN
		g_oFSO.DeleteFile sPictPath, TRUE
	END IF
	
END SUB


FUNCTION picture_saveFile( sField, sUploadFile, sPictName )

	DIM	sFileExt

	picture_saveFile = ""
	sFileExt = LCASE(g_Upload.ExtractFileExt( sUploadFile ))
	IF INSTR( 1, ".gif.jpg.png.", sFileExt&".", vbTextCompare ) THEN
	
		picture_saveFile = sPictName & sFileExt
		DIM	sPictSpec
		sPictSpec = g_oFSO.BuildPath( g_sPicturePath, picture_saveFile )
		picture_debug sPictSpec & "<br>"
		IF g_oFSO.FileExists( sPictSpec ) THEN g_oFSO.DeleteFile sPictSpec, TRUE
		ON ERROR Resume Next
		IF NOT g_Upload.SaveFile( sField, sPictSpec ) THEN
			picture_saveFile = ""
			picture_debug "<br><br>Error saving file<br><br>"
		END IF
		ON ERROR Goto 0
	END IF

END FUNCTION



%>
