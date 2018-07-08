<%@	Language=VBScript%>
<%
OPTION EXPLICIT


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%




SUB cleanDownloadFolder( sDownloadFolder, sPicturePath )
	DIM	sFileLC
	sFileLC = LCASE(sPicturePath)
	DIM	oFolder
	SET oFolder = g_oFSO.GetFolder( sDownloadFolder )
	IF NOT oFolder IS Nothing THEN
		DIM	oFile
		FOR EACH oFile IN oFolder.Files
			IF sFileLC <> LCASE(oFile.Path) THEN
				IF 5 < ABS(DATEDIFF("d",NOW,oFile.DateLastModified)) THEN
					oFile.Delete(TRUE)
				END IF
			END IF
		NEXT
	END IF
END SUB




DIM	nImageID
DIM	sImageID
sImageID = Request.QueryString("id")
IF ISNUMERIC(sImageID) THEN
	nImageID = CLNG(sImageID)
ELSE
	nImageID = 0
END IF

DIM sMime
DIM sFilename
DIM	bFound
bFound = FALSE
sMime = "*/*"
sFilename = ""

IF 0 < nImageID THEN
	DIM	sSelect
	sSelect = "" _
		&	"SELECT * " _
		&	"FROM images " _
		&	"WHERE RID=" & nImageID & ";"
	DIM	oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		sMime = recString(oRS.Fields("Mime"))
		Response.ContentType = sMime
		IF "image" = MID(sMime,1,5) THEN
			Response.BinaryWrite oRS.Fields("Data").Value
			bFound = TRUE
		ELSE
			sFilename = recString(oRS.Fields("File"))
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
END IF

DIM	oStream
IF NOT bFound THEN
	IF 0 < LEN(sFilename) THEN
		sPicturePath = picture_buildPath( "Pages", sFilename )
		IF g_oFSO.FileExists( sPicturePath ) THEN
			SET oStream = Server.CreateObject("ADODB.Stream")
			IF NOT oStream IS Nothing THEN
				oStream.Open
				oStream.Type = adTypeBinary
				oStream.LoadFromFile sPicturePath
				Response.ContentType = sMime
				Response.BinaryWrite oStream.Read
				oStream.Close
				bFound = TRUE
			END IF
			SET oStream = Nothing
		END IF
	END IF
END IF

IF NOT bFound THEN
	DIM	sPicturePath
	sPicturePath = Server.MapPath("images/notfound.jpg")
	IF g_oFSO.FileExists( sPicturePath ) THEN

		SET oStream = Server.CreateObject("ADODB.Stream")

		IF NOT oStream IS Nothing THEN

			oStream.Open
			oStream.Type = adTypeBinary

			oStream.LoadFromFile sPicturePath

			Response.ContentType = "image/jpg"
			Response.BinaryWrite oStream.Read
	
			oStream.Close
			bFound = TRUE

		END IF

		SET oStream = Nothing

	END IF
END IF

IF NOT bFound THEN
	Response.Status = "404 Not Found"
	Response.Write "404 Not Found"
END IF






IF FALSE THEN


sDownloadFolder = findFolder( "download" )
IF "" <> sDownloadFolder THEN
	DIM	sDownloadFile
	sDownloadFile = g_oFSO.BuildPath( sDownloadFolder, sName )
	cleanDownloadFolder sDownloadFolder, sPicturePath
	IF g_oFSO.FileExists( sDownloadFile ) THEN
		DIM	oFile
		SET oFile = g_oFSO.GetFile( sDownloadFile )
		IF NOT oFile IS Nothing THEN
			DIM	oFile1
			SET oFile1 = g_oFSO.GetFile( sPicturePath )
			IF 5 < ABS(DATEDIFF("d",NOW,oFile.DateLastModified))  OR  oFile.DateLastModified < oFile1.DateLastModified THEN
				oFile.Delete(TRUE)
				SET oFile = Nothing
				g_oFSO.CopyFile sPicturePath, sDownloadFile
				DIM oShell
				DIM oFolder
				Set oShell = CreateObject("Shell.Application")
				Set oFolder = oShell.NameSpace(sDownloadFolder)
				Set oFile = oFolder.ParseName(sName)
				oFile.ModifyDate = NOW 
				set oFile = nothing 
				set oFolder = nothing 
				set oShell = nothing 
			END IF
			SET oFile = Nothing
		END IF
	ELSE
		g_oFSO.CopyFile sPicturePath, sDownloadFile
	END IF
	DIM	sVirt
	sVirt = virtualFromPhysicalPath( sDownloadFile )
	sVirt = "http://" & Request.ServerVariables("HTTP_HOST") & sVirt
	Response.Redirect sVirt

ELSE

	DIM	sSuffix
	DIM	i
	i = INSTRREV( sPicturePath, ".", -1, vbTextCompare )
	IF 0 < i THEN
		sSuffix = MID( sPicturePath, i+1 )
	ELSE
		sSuffix = "unknown"
	END IF
	sLabel = sLabel & "." & sSuffix
	
	'DIM	oStream
	SET oStream = Server.CreateObject("ADODB.Stream")
	
	IF NOT oStream IS Nothing THEN
	
		oStream.Open
		oStream.Type = adTypeBinary
	
		oStream.LoadFromFile sPicturePath
		
		Response.AddHeader "content-disposition","attachment; filename=" & sLabel
	
		Response.ContentType = getContentType( sPicturePath )
		Response.BinaryWrite oStream.Read
		
		oStream.Close
	
		'Response.End
	
	END IF
	
	SET oStream = Nothing
END IF
SET g_oFSO = Nothing


END IF	' FALSE


%>
<!--#include file="scripts\include_db_end.asp"-->
