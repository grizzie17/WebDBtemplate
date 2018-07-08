<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Server.ScriptTimeout = 60*30
Response.Expires = 0
Response.Buffer = True



%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"--><%


DIM	sCategory
sCategory = Request("category")

DIM	sBrand
sBrand = Request("brand")





	DIM	sCategoryName
	sCategoryName = ""

IF "" <> sCategory THEN
	DIM	sSelect
				sSelect = "" _
					&	"SELECT " _
					&		dbQ("Name") & " " _
					&	"FROM " _
					&		"categories " _
					&	"WHERE " _
					&		"CategoryID = " & sCategory _
					&	";"

	DIM	oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT Nothing IS oRS THEN
		IF 0 < oRS.RecordCount THEN
			sCategoryName = oRS.Fields("Name").Value
			sCategoryName = Server.HTMLEncode(sCategoryName)
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
ELSE
	sCategoryName = "Undefined Category"
END IF

%><html>

<head>
<title>Add Page (<%=sCategoryName%>)</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="scripts/formvalidate.js"></script>
<script language="JavaScript" type="text/javascript" src="scripts/pageedit.js"></script>
<script language="JavaScript" type="text/javascript">
<!--

function doCancel()
{
	window.location.href = "pages.asp?category=<%=sCategory%>";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}


//-->
</script>
</head>

<body bgcolor="#FFFFFF" id="pageAdd">

<h1>Add Page (<%=sCategoryName%>)</h1>
<%

FUNCTION checkTrue( o )
	checkTrue = ""
	IF NOT ISNULL( o ) THEN
		IF CBOOL( o ) THEN
			checkTrue = "checked"
		ELSE
			checkTrue = ""
		END IF
	END IF
END FUNCTION






%>
<form method="POST" action="pageaddsubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
	<input type="hidden" name="QualCategory" value="<%=sCategory%>">
	<div align="center">
		<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
			<tr>
				<td align="right" valign="top">Category</td>
				<td><select size="1" name="Category" class="required"><%

DIM	sWhereCat
sWhereCat = ""
IF "" <> sCategory THEN
	sWhereCat = "OR CategoryID=" & sCategory & " "
END IF
	sSelect = "SELECT " _
			&		"CategoryID, " _
			&		"Name, " _
			&		"Disabled " _
			&	"FROM " _
			&		"categories " _
			&	"WHERE " _
			&		"0 = Disabled " _
			&		sWhereCat & " " _
			&	"ORDER BY " _
			&		"Name" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oCategoryID
	DIM oCatName
	DIM	oCatDisabled
	
	SET oCategoryID = oRS.Fields("CategoryID")
	SET oCatName = oRS.Fields("Name")
	SET oCatDisabled = oRS.Fields("Disabled")
	
	oRS.MoveFirst
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oCategoryID & """"
		IF "" <> sCategory THEN
			IF recNumber(oCategoryID) = CLNG(sCategory) THEN
				Response.Write " selected"
				IF 0 <> oCatDisabled THEN
					Response.Write " style=""color: #CCCCCC"""
				END IF
			END IF
		END IF
		Response.Write ">" & Server.HTMLEncode(oCatName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
	
	oRS.Close
	SET oRS = Nothing
  %></select></td>
			</tr>
			<tr>
				<td align="right">Name</td>
				<td>
				<input type="text" name="PageName" size="40" maxlength="64" value="" class="required"></td>
			</tr>
			<tr>
				<td align="right">Sort Name</td>
				<td><input type="text" name="SortName" size="20" maxlength="32" value=""></td>
			</tr>
			<tr>
				<td align="right">Title</td>
				<td>
				<input type="text" name="Title" id="Title" size="50" maxlength="128" value="" class="required">
				<a href="javascript:fixupTextInput('Title')">fix-special-chars</a></td>
			</tr>
			<tr>
				<td align="right">Description</td>
				<td>
				<input type="text" name="Description" id="Description" size="50" maxlength="255" value="">
				<a href="javascript:fixupTextInput('Description')">fix-special-chars</a></td>
			</tr>
		</table>
		<p><input type="submit" value="Add" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>
<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->