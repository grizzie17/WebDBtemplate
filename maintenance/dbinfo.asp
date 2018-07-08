<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!DOCTYPE html>
<html>
	<head>
		<meta charset="UTF-8"/>
		<title>DB Info</title>
	</head>
	<body>
<%
	Response.Write "<p>Cookie-Prefix=" & g_sCookiePrefix & "</p>" & vbCRLF
	Response.Write "<p>DSN=grizzie_" & g_sDBSource  & "</p>" & vbCRLF
%>
<!--#include file="scripts\include_db_begin_admin.asp"-->
<!--#include file="scripts\include_db_end.asp"-->
	</body>
</html>