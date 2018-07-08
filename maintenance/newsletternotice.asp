<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<title>Newsletter Notification</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
</head>

<body>

<%

DIM	dToday
dToday = NOW
DIM	dd
DIM	mm
DIM	yy
dd = DAY(dToday)
mm = MONTH(dToday)

IF 20 < dd THEN mm = mm + 1
IF 12 < mm THEN mm = 1

%>

<h2>Send Email Notice of Newsletter Availability</h2>
<form method="post" action="newsletternoticesend.asp">
<div style="position: relative">
<div style="position: relative; display: block; margin-right: auto; margin-left: auto; font-family: Arial, Helvetica, sans-serif;">
<p><select name="Month">
<%
DIM	i
FOR i = 1 TO 12
	Response.Write "<option value=""" & i & """"
	IF i = mm THEN
		Response.Write " selected"
	END IF
	Response.Write ">" & MONTHNAME(i) & "</option>" & vbCRLF
NEXT
%>
</select></p>
<p>Additional Comments:<br>
<textarea name="Comments" cols="40" rows="10"></textarea></p>
<p><input name="Submit" type="submit" value="Send Email"></p>
</div>
</div>
</form>
<!--#include file="scripts\byline.asp"-->

</body>

</html>
