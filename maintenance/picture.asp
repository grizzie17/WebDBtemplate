<%@	Language=VBScript%>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\ado.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%


FUNCTION getMime( s )

	getMime = "*/*"
	DIM	sSuffix
	DIM	aSuffix
	DIM	aMime
	DIM	i
	DIM	oFile
	DIM	sFile
	DIM	sData
	
	sSuffix = LCASE( s )
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

FUNCTION getContentType( s )

	DIM	i
	DIM	sSuffix
	i = INSTRREV( s, ".", -1, vbTextCompare )
	IF 0 < i THEN
		sSuffix = MID( s, i+1 )
		getContentType = getMime( sSuffix )
	ELSE
		getContentType = "*/*"
	END IF
	
END FUNCTION



DIM	sTable
sTable = Request.QueryString("table")

DIM	sName
sName = Request.QueryString("name")

DIM	sPicturePath

IF "" <> sTable  AND  "" <> sName THEN

	sPicturePath = picture_buildPath( sTable, sName )
	IF NOT g_oFSO.FileExists( sPicturePath ) THEN
		sPicturePath = Server.MapPath("images/notfound.jpg")
	END IF

ELSE
	sPicturePath = Server.MapPath("images/notfound.jpg")
END IF

DIM	oStream
SET oStream = Server.CreateObject("ADODB.Stream")

IF NOT oStream IS Nothing THEN

	oStream.Open
	oStream.Type = adTypeBinary

	oStream.LoadFromFile sPicturePath

	Response.ContentType = getContentType( sPicturePath )
	Response.BinaryWrite oStream.Read
	
	oStream.Close

	'Response.End

END IF

SET oStream = Nothing
SET g_oFSO = Nothing


%>