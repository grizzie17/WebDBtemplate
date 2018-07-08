<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


%>
<!--#include virtual="/scripts/config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include virtual="/scripts/include_db_begin.asp"-->
<%




%><html>

<head>
<title>Add Category</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script type="text/javascript" language="JavaScript" src="scripts/formvalidate.js"></script>
<script type="text/javascript" language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "categories.asp";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}


//-->
</script>
</head>

<body bgcolor="#FFFFFF" id="pageAdd">

<h1>Add Category</h1>
<%



DIM	sSelect
DIM	oRS

DIM	nMxsName
DIM	nMxsSortName
DIM	nMxsShortName

nMxsName = dbGetFieldSize( g_DC, "categories", "Name")
nMxsSortName = dbGetFieldSize( g_DC, "categories", "SortName")
nMxsShortName = dbGetFieldSize( g_DC, "categories", "ShortName")




%>
<form method="POST" action="categoryaddsubmit.asp" name="EditForm" onsubmit="return validateForm()">
	<div align="center">
		<table border="0">
			<tr>
				<td valign="top">
				<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
					<tr>
						<td align="right">Name</td>
						<td>
						<input type="text" name="Name" size="40" maxlength="<%=nMxsName%>" value class="required"></td>
					</tr>
					<tr>
						<td align="right">Sort Name</td>
						<td>
						<input type="text" name="SortName" size="20" maxlength="<%=nMxsSortName%>" value></td>
					</tr>
					<tr>
						<td align="right">Short Name</td>
						<td>
						<input type="text" name="ShortName" size="30" maxlength="<%=nMxsShortName%>" value>
						</td>
					</tr>
					<tr>
						<td align="right" valign="top">Description</td>
						<td><textarea rows="7" name="ShortDesc" cols="35"></textarea></td>
					</tr>
				</table>
				</td>
				<td valign="top"><%

'DIM	oRS
'DIM	sSelect

sSelect = "" _
	&	"SELECT " _
	&		"TabID, " _
	&		"Name " _
	&	"FROM " _
	&		"tabs " _
	&	"WHERE " _
	&		"0 = Disabled " _
	&	"ORDER BY " _
	&		"Name " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )

IF 0 < oRS.RecordCount THEN

	DIM	oID
	DIM	oName
	
	SET oID = oRS.Fields("TabID")
	SET oName = oRS.Fields("Name")

%>
				<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
					<tr>
						<td><%


%> Tabs<br>
						<select size="7" name="Tabs" multiple><%
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oID & """>" & Server.HTMLEncode(oName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
%></select> <%

%> </td>
					</tr>
				</table>
				<%

END IF

oRS.Close

SET oRS = Nothing

%> </td>
			</tr>
		</table>
		<p><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>

<!--#include virtual="/scripts/byline.asp"-->

</body>

</html>
<!--#include virtual="/scripts/include_db_end.asp"-->