<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Server.ScriptTimeout = 60*30
Session.Timeout = 90
Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon_defs.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\vbscriptex.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">
<meta content="en-us" http-equiv="Content-Language">
<title>Logon</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
</head>

<body id="pageLogon">

<%

DIM	sError
sError = Request("error")
IF "" <> sError THEN
	IF ISNUMERIC(sError) THEN
		sError = authErrorString( CLNG(sError) )
	END IF
	Response.Write "<p style=""color:red"">" & Server.HTMLEncode(sError) & "</p>" & vbCRLF
END IF


DIM	sSelect
DIM	oRS

DIM	nMxsUsername
DIM	nMxsPassword

nMxsUsername = dbGetFieldSize( g_DC, "authorizeusers", "UserName" )
nMxsPassword = dbGetFieldSize( g_DC, "authorizeusers", "Password" )





%>

<h1>Logon</h1>
<form action="logonsubmit.asp" method="post">
<div style="position: relative; margin-right: auto; margin-left: auto; font-family: Arial, Helvetica, sans-serif;">
<table style="position: relative; margin-right: auto; margin-left: auto">
	<tr>
		<td>Username:</td>
		<td><input name="Username" size="<%=min(24,nMxsUsername)%>" maxlength="<%=nMxsUsername%>" type="text"></td>
	</tr>
	<tr>
		<td>Password:</td>
		<td><input name="Password" size="<%=min(24,nMxsPassword)%>" maxlength="<%=nMxsPassword%>" type="password"></td>
	</tr>
<%
IF FALSE THEN
%>
	<tr>
		<td>&nbsp;</td>
		<td><input name="Remember" id="Remember" type="checkbox" checked="checked" value="YES"><label id="Label1" for="Remember"> Remember my logon 
		for this computer (should not be used if shared)</label>
		<div class="instructions">Note: Remember expires if you do not connect again within 14 days</div></td>
	</tr>
<%
END IF
%>
	<tr>
		<td>&nbsp;</td>
		<td><input name="Submit1" type="submit" value="Logon"></td>
	</tr>
</table>
</div>

</form>
<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
