<!--#include file="upload3\code\clsUpload.asp"-->
<!--#include file="upload3\code\clsImage.asp"-->
<%

DIM	g_sPictureSuffixes
g_sPIctureSuffixes = ".gif.jpg.png."


DIM	g_sPicturePath
g_sPicturePath = ""

DIM g_Upload
SET g_Upload = Nothing

DIM	g_sUploadClass
g_sUploadClass = ""

DIM	g_bUseRequestObject
g_bUseRequestObject = FALSE


SUB picture_createObject()
	SET g_Upload = Nothing
	ON ERROR Resume Next
	
	IF g_bUseRequestObject THEN
		SET g_Upload = Request
		g_sUploadClass = "Request"
	END IF
	
	IF g_Upload IS Nothing THEN
		SET g_Upload = Server.CreateObject("Persits.Upload")
		IF NOT g_Upload IS Nothing THEN
			g_sUploadClass = "PersitsUpload"
			g_Upload.Save
		END IF
	END IF
	
	IF g_Upload IS Nothing THEN
		SET g_Upload = Server.CreateObject("ASPSimpleUpload.Upload")
		IF NOT g_Upload IS Nothing THEN
			g_sUploadClass = "ASPSimpleUpload"
		END IF
	END IF
	
	IF g_Upload IS Nothing THEN
		SET g_Upload = NEW clsUpload
		IF NOT g_Upload IS Nothing THEN
			g_sUploadClass = "clsUpload"
		END IF
	END IF
	ON ERROR Goto 0
	IF Nothing IS g_Upload THEN
		g_sUploadClass = ""
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
	SELECT CASE LCASE(g_sUploadClass)
	CASE "persitsupload"
		Requestor = g_Upload.Form( s )
	CASE "aspsimpleupload"
		Requestor = g_Upload.Form( s )
	CASE "clsupload"
		DIM	oCol
		DIM	oItem
		DIM	sTemp
		oCol = g_Upload.Collection( s )
		sTemp = ""
		FOR EACH oItem IN oCol
			IF RIGHT(oItem.Value,2) = vbCRLF THEN
				sTemp = sTemp & "," & LEFT(oItem.Value,LEN(oItem.Value)-2)
			ELSE
				sTemp = sTemp & "," & oItem.Value
			END IF
		NEXT 'oItem
		Requestor = MID(sTemp,2)
	CASE ELSE
		Requestor = Request( s )
		picture_debug " : "
	END SELECT
	picture_debug " = "
	picture_debug Requestor & "<br>"
END FUNCTION



FUNCTION RequestorFile( s )
	picture_debug s
	SELECT CASE LCASE(g_sUploadClass)
	CASE "persitsupload"
		DIM	oUF
		SET oUF = g_Upload.Files( s )
		IF NOT oUF IS Nothing THEN
			RequestorFile = oUF.OriginalFileName
			SET oUF = Nothing
		ELSE
			RequestorFile = ""
		END IF
	CASE "aspsimpleupload"
		RequestorFile = g_Upload.Form( s )
	CASE "clsupload"
		RequestorFile = Requestor( s )
	CASE ELSE
		RequestorFile = Request( s )
		picture_debug " : "
	END SELECT
	picture_debug " = " & RequestorFile & "<br>"
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


FUNCTION picture_buildPath( sTable, sPictFile )
	picture_buildPath = ""
	IF "" = g_sPicturePath THEN
		picture_findPath sTable
	END IF
	IF "" <> g_sPicturePath THEN
		picture_buildPath = g_oFSO.BuildPath( g_sPicturePath, sPictFile )
	END IF
END FUNCTION


SUB picture_delete( sFile )

	DIM	sPictPath
	
	sPictPath = g_oFSO.BuildPath( g_sPicturePath, sFile )
	IF g_oFSO.FileExists( sPictPath ) THEN
		g_oFSO.DeleteFile sPictPath, TRUE
	END IF
	
END SUB


SUB picture_resize( sPictSpec, nMaxWidth, nMaxHeight )

	IF 0 < nMaxWidth  OR  0 < nMaxHeight THEN
		DIM	oJPEG
		SET oJPEG = Nothing
		ON ERROR Resume Next
		SET oJPEG = Server.CreateObject("Persits.Jpeg")
		ON ERROR Goto 0
		IF NOT oJPEG IS Nothing THEN
			oJPEG.Open sPictSpec
			
			DIM h
			DIM	w
			DIM	hRatio
			DIM	wRatio
			DIM	fRatio
			DIM	bScale
			
			h = oJPEG.OriginalHeight
			w = oJPEG.OriginalWidth
			bScale = FALSE
			
			IF 0 < nMaxWidth  AND  0 < nMaxHeight THEN
				IF 0 < w  AND  0 < h THEN
					IF nMaxWidth < w  OR  nMaxHeight < h THEN
						wRatio = w / nMaxWidth
						hRatio = h / nMaxWidth
						IF hRatio < wRatio THEN
							fRatio = wRatio
						ELSE
							fRatio = hRatio
						END IF
						w = FIX(w / fRatio)
						h = FIX(h / fRatio)
						bScale = TRUE
					END IF
				END IF
			ELSEIF 0 < nMaxWidth  AND  nMaxWidth < w THEN
				bScale = TRUE
				h = h * nMaxWidth / w
				w = nMaxWidth
			ELSEIF 0 < nMaxHeight  AND  nMaxHeight < h THEN
				bScale = TRUE
				w = w * nMaxHeight / h
				h = nMaxHeight
			END IF
			IF bScale THEN
				'oJPEG.PreserveAspectRatio = TRUE
				oJPEG.Width = w
				oJPEG.Height = h
				IF g_oFSO.FileExists( sPictSpec ) THEN
					g_oFSO.DeleteFile sPictSpec, TRUE
				END IF
				oJPEG.Save sPictSpec
			END IF
			oJPEG.Close
		END IF
		SET oJPEG = Nothing
	END IF
END SUB





FUNCTION picture_saveFile( sField, sUploadFile, sPictName, nMaxWidth, nMaxHeight )

	DIM	sFileExt
	DIM	sPictSpec
	DIM	sUFilename
	DIM	i
	
	SELECT CASE LCASE(g_sUploadClass)
	CASE "persitsupload"
	
		picture_saveFile = ""
		DIM	oUFile
		SET oUFile = g_Upload.Files(sField)
		IF NOT oUFile IS Nothing THEN
			sUFilename = oUFile.Filename
			i = INSTRREV(sUFilename,".")
			IF 0 < i THEN
				sFileExt = MID(sUfilename,i)
				IF 0 < INSTR(g_sPictureSuffixes, LCASE(sFileExt)&".") THEN
					picture_saveFile = sPictName & sFileExt
					sPictSpec = g_oFSO.BuildPath( g_sPicturePath, picture_saveFile )
					picture_debug sPictSpec & "<br>"
					IF g_oFSO.FileExists( sPictSpec ) THEN g_oFSO.DeleteFile sPictSpec, TRUE
					ON ERROR Resume Next
					oUFile.SaveAs sPictSpec
					ON ERROR Goto 0
				END IF
			END IF
			SET oUFile = Nothing
		END IF
		
	CASE "aspsimpleupload"
	
		picture_saveFile = ""
		sFileExt = LCASE(g_Upload.ExtractFileExt( sUploadFile ))
		IF 0 < INSTR( g_sPictureSuffixes, LCASE(sFileExt)&"." ) THEN
		
			picture_saveFile = sPictName & sFileExt
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
		
	CASE "clsupload"
	
		picture_saveFile = ""
		sUFilename = g_Upload.Fields(sField).Filename
		i = INSTRREV(sUFilename,".")
		IF 0 < i THEN
			sFileExt = MID(sUFilename,i)
			IF 0 < INSTR(g_sPictureSuffixes, LCASE(sFileExt)&".") THEN
				picture_saveFile = sPictName & sFileExt
				sPictSpec = g_oFSO.BuildPath( g_sPicturePath, picture_saveFile )
				picture_debug sPictSpec & "<br>"
				IF g_oFSO.FileExists( sPictSpec ) THEN g_oFSO.DeleteFile sPictSpec, TRUE
				ON ERROR Resume Next
				g_Upload.Fields(sField).SaveAs sPictSpec
				ON ERROR Goto 0
			END IF
		END IF
	
	CASE ELSE
	END SELECT
	
	IF "" <> picture_saveFile THEN
		IF "" <> sPictSpec THEN
			IF "" <> sFileExt THEN
				IF 0 < INSTR(".jpg.gif.png.", LCASE(sFileExt)&".") THEN
					picture_resize sPictSpec, nMaxWidth, nMaxHeight
				END IF
			END IF
		END IF
	END IF


END FUNCTION


FUNCTION getMime( s )

	getMime = "*/*"
	DIM	sSuffix
	DIM	aSuffix
	DIM	aMime
	DIM	i
	DIM	oFile
	DIM	sFile
	DIM	sData

	i = INSTRREV( s, ".", -1, vbTextCompare )
	IF 0 < i THEN
		sSuffix = LCASE(MID( s, i+1 ))
	ELSE
		sSuffix = LCASE( s )
	END IF
	sFile = findFileUpTree("scripts\mime.txt")
	IF "" <> sFile THEN
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT oFile IS Nothing THEN
			sData = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			aMime = SPLIT(sData, vbCRLF )
			DIM	sLine
			FOR EACH sLine IN aMime
				IF "" <> sLine THEN
					aSuffix = SPLIT(sLine,vbTAB)
					IF LBOUND(aSuffix) < UBOUND(aSuffix) THEN
						IF sSuffix = aSuffix(0) THEN
							getMime = aSuffix(1)
							EXIT FOR
						END IF
					END IF
				END IF
			NEXT
		END IF
	END IF

END FUNCTION


FUNCTION storeImage( oDC, sPictureFile )
	'Response.Write "storeImage( " & sPictureFile & ")<br>" & vbCRLF
	storeImage = 0
	DIM	oStream
	SET oStream = Server.CreateObject( "ADODB.Stream" )
	IF NOT oStream IS Nothing THEN
		oStream.Type = adTypeBinary
		oStream.Open
		oStream.LoadFromFile sPictureFile

		DIM	i
		DIM	sFileName
		i = INSTRREV( sPictureFile, "\" )
		IF 0 < i THEN
			sFileName = MID(sPictureFile, i+1)
		ELSE
			sFileName = sPictureFile
		END IF

		DIM	sSelect
		sSelect = "" _
			&	"SELECT * " _
			&	"FROM " & dbQ("images") & " " _
			&	"WHERE RID=0;"
		DIM	oRS
		SET oRS = dbQueryUpdate( oDC, sSelect )
		IF NOT oRS IS Nothing THEN
			oRS.AddNew
			oRS.Fields("File").Value = sFileName
			oRS.Fields("Mime").Value = getMime( sFileName )
			oRS.Fields("DateCreated").Value = NOW()
			oRS.Fields("Data").Value = oStream.Read
			oRS.Update
			storeImage = oRS.Fields("RID").Value
			oRS.Close
			SET oRS = Nothing
		END IF
		oStream.Close
		SET oStream = Nothing
	END IF
	g_oFSO.DeleteFile sPictureFile
END FUNCTION


%>
