<%@	Language=VBScript%>
<%
OPTION EXPLICIT



%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<%





DIM	nImageID
DIM	sImageID
sImageID = Request.QueryString("id")
IF ISNUMERIC(sImageID) THEN
	nImageID = CLNG(sImageID)
ELSE
	nImageID = 0
END IF

DIM	bFound
bFound = FALSE

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
		Response.ContentType = recString(oRS.Fields("Mime"))
		Response.BinaryWrite oRS.Fields("Data").Value
		oRS.Close
		SET oRS = Nothing
		bFound = TRUE
	END IF
END IF

IF NOT bFound THEN
	DIM	sPicturePath
	sPicturePath = Server.MapPath("images/notfound.jpg")
	IF g_oFSO.FileExists( sPicturePath ) THEN

		DIM	oStream
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



%>
<!--#include file="scripts\include_db_end.asp"-->
