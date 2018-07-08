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


DIM	sCategoryID
sCategoryID = Request("category")




%><html>

<head>
<title>Edit Category</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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

<body bgcolor="#FFFFFF" id="pageEdit">

<h1>Edit Category</h1>
<%

FUNCTION checkTrue( o )
	checkTrue = ""
	IF recBool( o ) THEN
		checkTrue = "checked"
	END IF
END FUNCTION





	DIM	sSelect
	sSelect = "SELECT " _
			&	" * " _
			&	"FROM categories " _
			&	"WHERE " _
			&		"CategoryID=" & sCategoryID & " " _
			&	"ORDER BY " _
			&		"Name " _
			&	";"
	
	
	DIM	RS
	SET RS = dbQueryRead( g_DC, sSelect )



	IF RS.EOF Then
		RS.Close
		Response.Write "<p>Requested Category not found</p>"
	ELSE
		DIM	oRID
		DIM	oPG
		DIM	oCategoryID
		DIM	oName
		DIM	oSortName
		DIM	oShortName
		DIM	oShortDesc
		DIM	oDisabled
				
		SET oCategoryID = RS.Fields("CategoryID")
		SET	oName = RS.Fields("Name")
		SET oSortName = RS.Fields("SortName")
		SET	oShortName = RS.Fields("ShortName")
		SET oShortDesc = RS.Fields("ShortDesc")
		SET	oDisabled = RS.Fields("Disabled")
		
		DIM	nMxsName
		DIM	nMxsSortName
		DIM	nMxsShortName
		
		nMxsName = oName.DefinedSize
		nMxsSortName = oSortName.DefinedSize
		nMxsShortName = oShortName.DefinedSize


%>

<form method="POST" action="categoryeditsubmit.asp" name="EditForm" onsubmit="return validateForm()">
<input type="hidden" name="CategoryID" value="<%=recNumber(oCategoryID)%>">
<div align="center">
<table border="0">
<tr>
<td valign="top">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
          <td align="right">Name</td>
          <td>
  <input type="text" name="Name" size="40" maxlength="<%=nMxsName%>" value="<%=recString(oName)%>" class="required"></td>
        </tr>
    <tr>
          <td align="right">Sort Name</td>
          <td>
  <input type="text" name="SortName" size="20" maxlength="<%=nMxsSortName%>" value="<%=Server.HTMLEncode(recString(oSortName))%>"></td>
        </tr>
    <tr>
          <td align="right">Short Name</td>
          <td><input type="text" name="ShortName" size="30" maxlength="<%=nMxsShortName%>" value="<%=recString(oShortName)%>">
  </td>
        </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td>
  <textarea rows="7" name="ShortDesc" cols="35"><%=recString(oShortDesc)%></textarea></td>
  </tr>
  <tr>
  <td></td>
      <td>
  <input type="checkbox" name="Disabled" value="ON" <%=checkTrue(oDisabled)%>>Disabled</td>
    </tr>
  </table>
</td>
<td valign="top">
<%


DIM	oRS
'DIM	sSelect

sSelect = "" _
	&	"SELECT " _
	&		"tabs.TabID AS TabID, " _
	&		"tabs.Name AS Name, " _
	&		"tabs.Disabled AS TabDisabled, " _
	&		"tabcategorymap.CategoryID AS CategoryID " _
	&	"FROM " _
	&		"tabs " _
	&		"LEFT JOIN tabcategorymap " _
	&			"ON (tabcategorymap.TabID = Tabs.TabID " _
	&				"AND tabcategorymap.CategoryID = " & sCategoryID & ") " _
	&	"WHERE " _
	&		"0 = Tabs.Disabled " _
	&		"OR tabcategorymap.CategoryID = " & sCategoryID & " " _
	&	"ORDER BY " _
	&		"Name " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )

IF 0 < oRS.RecordCount THEN

	DIM	oID
	DIM	oTab
	DIM	oTabDisabled
	'DIM	oName
	
	SET oID = oRS.Fields("TabID")
	SET oName = oRS.Fields("Name")
	SET oTab = oRS.Fields("CategoryID")
	SET	oTabDisabled = oRS.Fields("TabDisabled")

%>
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<td>
<%


%>
Tabs<br>
<select size="7" name="Tabs" multiple>
<%
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oID & """"
		IF NOT ISNULL(oTab) THEN Response.Write " selected"
		IF isTrue(oTabDisabled) THEN Response.Write " style=""color: #CCCCCC"""
		Response.Write ">" & Server.HTMLEncode(oName) & "</option>" & vbCRLF
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
'SET oDC = Nothing

	RS.Close
	SET RS = Nothing
	END IF

%>
</td>
</tr>
</table>
  <p>
  <input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
  <input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
</div>

</form>


<!--#include file="scripts\byline.asp"-->

</body>

</html>
<%
SET	RS = Nothing
%><!--#include file="scripts\include_db_end.asp"-->