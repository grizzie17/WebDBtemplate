<%@ Language=VBScript %>
<%
OPTION EXPLICIT



DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="config.asp"-->
<%

DIM	sFile
DIM	sTester
DIM	sResult
sResult = ""
sTester = Request.Cookies(g_sCookiePrefix & "_CookiesTester")
IF "" <> sTester THEN
	sResult = Request.Cookies(sTester)
END IF

IF 5 < LEN(sResult) THEN
	sFile = "pobox.js"
ELSE
	sFile = "pobox_nospam.js"
END IF

DIM	sJS

sJS = Server.MapPath(sFile)
IF g_oFSO.FileExists(sJS) THEN

	DIM	oFile
	SET oFile = g_oFSO.OpenTextFile( sJS, 1 )
	IF NOT oFile IS Nothing THEN
		
		DIM	sContent
		sContent = oFile.ReadAll
		
		Response.ContentType = "text/javascript"
		Response.Write sContent
		
		oFile.Close
		SET oFile = Nothing
	END IF
ELSE

	Response.ContentType = "text/javascript"
	Response.Write ""
END IF



SET g_oFSO = Nothing

%>