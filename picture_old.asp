<%@ Language=VBScript %>
<%
OPTION EXPLICIT

Response.Expires = 30


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\ado.inc"-->
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_picture_old.asp"-->
<%



DIM	sTable
sTable = Request.QueryString("file")

DIM	q_h
DIM	q_w
q_h = Request.QueryString("h")
q_w = Request.QueryString("w")

DIM	q_copy
q_copy = Request.QueryString("c")



FUNCTION IsNum( s )
	IsNum = FALSE
	IF "" <> CSTR(s) THEN
		IF ISNUMERIC(s) THEN IsNum = TRUE
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
		IF "asp" = sSuffix  OR  "xslt" = sSuffix THEN
			getContentType = ""
		END IF
	ELSE
		getContentType = "image/gif"
	END IF
	
END FUNCTION





DIM	sName

DIM	sPicturePath

IF "" <> sTable THEN

	IF ":" = MID(sTable,2,1) THEN
		sPicturePath = sTable
	ELSE
		sPicturePath = Server.MapPath(sTable)
	END IF
	IF NOT g_oFSO.FileExists( sPicturePath ) THEN
		sPicturePath = Server.MapPath("images/notfound.jpg")
	END IF

ELSE
	sPicturePath = Server.MapPath("images/notfound.jpg")
END IF

DIM	cType
cType = getContentType( sPicturePath )
IF "" <> cType THEN

	DIM	oStream
	SET oStream = Nothing
	DIM	oCanvas
	SET oCanvas = Nothing
	IF 0 < INSTR(cType,"jpeg")  _
			AND  INSTR(sPicturePath,"notfound") < 1 _
			AND  INSTR(sPicturePath,"_thumbs") < 1  THEN
		ON ERROR Resume Next
		SET oCanvas = Server.CreateObject("Persits.Jpeg")
		ON ERROR GOTO 0
	END IF
	
	IF NOT oCanvas IS Nothing THEN
	
		oCanvas.Open sPicturePath
		oCanvas.Canvas.Font.Family = "Arial"
		
		DIM	h, w
		h = oCanvas.OriginalHeight
		w = oCanvas.OriginalWidth
		
		DIM	sCopy
		IF 0 < LEN(q_copy) THEN
			sCopy = REPLACE(q_copy,"(c)",CHRW(&HA9))
		ELSE
			sCopy = "Copyright " & CHRW(&HA9) & " " & YEAR(NOW) & " " & g_sSiteName
			sCopy = ""
		END IF
		
		IF IsNum(q_h)  AND  IsNum(q_w) THEN
			DIM	wRatio
			DIM	hRatio
			DIM	fRatio
			wRatio = w / CDBL(q_w)
			hRatio = h / CDBL(q_h)
			IF hRatio < wRatio THEN
				fRatio = wRatio
			ELSE
				fRatio = hRatio
			END IF
			w = FIX(w / fRatio)
			h = FIX(h / fRatio)
			oCanvas.Width = w
			oCanvas.Height = h
			oCanvas.Canvas.Font.Bold = 1
			oCanvas.Canvas.Font.Size = 16
			IF 0 < LEN(sCopy) THEN
				DIM	x, y
				x = 5
				y = h - (5+16)
				oCanvas.Canvas.Font.Color = &H333333
				oCanvas.Canvas.Print x, y+1, sCopy
				oCanvas.Canvas.Print x+1, y, sCopy
				oCanvas.Canvas.Font.Color = &H333333
				oCanvas.Canvas.Print x, y-1, sCopy
				oCanvas.Canvas.Print x-1, y, sCopy
				oCanvas.Canvas.Font.Color = &HFFFFFF
				oCanvas.Canvas.Print x, y, sCopy
			END IF
			
		ELSE
			IF 0 < LEN(sCopy) THEN
				oCanvas.Canvas.Font.Color = &HFFCC66
				oCanvas.Canvas.Font.Size = 24
				oCanvas.Canvas.Print 5, h-(5+24), sCopy
			END IF
		END IF
		
		Response.ContentType = cType
		Response.BinaryWrite oCanvas.Binary
		
		SET oCanvas = Nothing
		
	ELSE

		SET oStream = Server.CreateObject("ADODB.Stream")
		
		IF NOT oStream IS Nothing THEN
	
			oStream.Open
			oStream.Type = adTypeBinary
		
			oStream.LoadFromFile sPicturePath
		
			Response.ContentType = cType
			Response.BinaryWrite oStream.Read
			
			oStream.Close
			SET oStream = Nothing
		
		END IF
	END IF

	Response.End

END IF

SET g_oFSO = Nothing


%>
