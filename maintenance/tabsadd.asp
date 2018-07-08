<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<%




%><html>

<head>
<title>Add Tab</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="scripts/formvalidate.js"></script>
<script language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "tabs.asp";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}


//-->
</script>
</head>

<body bgcolor="#FFFFFF" id="pageAdd">

<h1>Add Tab</h1>
<%






%>

<form method="POST" action="tabsaddsubmit.asp" name="EditForm" onsubmit="return validateForm()">
<div align="center">
<table border="0">
<tr>
<td valign="top">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
          <td align="right">Name</td>
          <td>
  <input type="text" name="Name" size="40" maxlength="40" value="" class="required"></td>
        </tr>
    <tr>
          <td align="right">Sort Name</td>
          <td>
  <input type="text" name="SortName" size="20" maxlength="20" value=""></td>
        </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td><textarea rows="7" name="Description" cols="35"></textarea></td>
  </tr>
  </table>
</td>
<td valign="top">
<%

DIM	oRS
DIM	sSelect

sSelect = "" _
	&	"SELECT " _
	&		"CategoryID, " _
	&		"Name " _
	&	"FROM " _
	&		"categories " _
	&	"WHERE " _
	&		"0 = Disabled " _
	&	"ORDER BY " _
	&		"Name " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )

IF 0 < oRS.RecordCount THEN

	DIM	oID
	DIM	oName
	
	SET oID = oRS.Fields("CategoryID")
	SET oName = oRS.Fields("Name")

%>
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<td>
<%


%>
Categories<br>
<select size="7" name="Categories" multiple class="required">
<%
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oID & """>" & Server.HTMLEncode(oName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
%>
</select>
<%

%>
</td>
</tr>
</table>
<%

END IF

oRS.Close

SET oRS = Nothing

%>
</td>
</tr>
</table>
  <p>
  <input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
  <input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
</div>

</form>

<!--webbot bot="Include" U-Include="../_private/byline.htm" TAG="BODY" startspan -->


<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->