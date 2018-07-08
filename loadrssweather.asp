<%@ Language=VBScript %>
<%
OPTION EXPLICIT

DIM	g_oFSO
SET	g_oFSO = Server.CreateObject("Scripting.FileSystemObject")
%>
<!--#include file="scripts/findfiles.asp"-->
<!--#include file="scripts/rss.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<!--#include file="scripts/rssweather.asp"-->
<%



DIM	q_sZIP

q_sZIP = Request("z")
IF "" <> q_sZIP THEN
	DIM	sData
	sData = rssweather( q_sZIP )
	IF 0 < LEN(sData) THEN
		Response.ContentType = "text/text"
		Response.Write sData
	ELSE
		Response.Status = "404 Not Found"
	END IF
ELSE
	Response.Status = "403 Forbidden"
	Response.Write "403 Forbidden" & vbCRLF
END IF

Response.Flush
Response.End



SET g_oFSO = Nothing




%>