<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\htmlformat.asp"-->
<%
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">
<title>Add - Gallery</title>
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script type="text/javascript" language="JavaScript" src="scripts/formvalidate.js"></script>
<script type="text/javascript" language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "galleries.asp";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}


//-->
</script>
</head>

<body id="pageAdd">
<%

DIM	sSelect
sSelect = "" _
	&	"SELECT " _
	&		"* " _
	&	"FROM " _
	&		"Gallery " _
	&	"WHERE " _
	&		"RID = 0 " _
	&	";"

DIM	oRS
SET oRS = dbQueryRead( g_DC, sSelect )

DIM	nMxsTitle

nMxsTitle = oRS.Fields("Title").DefinedSize

oRS.Close
SET oRS = Nothing

%>
<form method="POST" action="galleryaddsubmit.asp" name="EditForm" onsubmit="return validateForm()">
	<div align="center">
<table border="0">
	<tr>
		<td align="right" valign="top">Title:</td>
		<td><input name="Title" size="50" maxlength="<%=nMxsTitle%>" type="text" class="required"></td>
	</tr>
	<tr>
		<td align="right" valign="top">Event Date:</td>
		<td><input name="EventDate" size="12" type="text" class="required"></td>
	</tr>
	<tr>
		<td align="right" valign="top">Description:</td>
		<td><textarea name="Description" cols="50" rows="15"></textarea></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="submit" name="save" value="Save"></td>
	</tr>
</table>
</div>
</form>
<%

%>

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
