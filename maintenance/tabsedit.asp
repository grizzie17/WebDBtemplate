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


DIM	sTabID
sTabID = Request("tab")




%><html>

<head>
<title>Edit Tab</title>
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

<body bgcolor="#FFFFFF" id="pageEdit">

<h1>Edit Tab</h1>
<%

FUNCTION isTrue( o )
	isTrue = ""
	IF NOT ISNULL( o ) THEN
		IF CBOOL( o ) THEN
			isTrue = "checked"
		ELSE
			isTrue = ""
		END IF
	END IF
END FUNCTION





	DIM	sSelect
	sSelect = "SELECT " _
			&	" tabs.*, " _
			&	" pagelistmap.ListID as TabSortID " _
			&	"FROM " _
			&		"tabs " _
			&		"LEFT JOIN pagelistmap ON (tabs.TabID = pagelistmap.TabID) " _
			&	"WHERE " _
			&		"tabs.TabID =" & sTabID & " " _
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
		DIM	oShortDesc
		DIM	oPicture
		DIM	oDisabled
		DIM	oTabSortID
				
		SET oCategoryID = RS.Fields("TabID")
		SET	oName = RS.Fields("Name")
		SET oSortName = RS.Fields("SortName")
		SET oShortDesc = RS.Fields("Description")
		SET oPicture = RS.Fields("Picture")
		SET	oDisabled = RS.Fields("Disabled")
		SET oTabSortID = RS.Fields("TabSortID")


%>

<form method="POST" action="tabseditsubmit.asp" name="EditForm" onsubmit="return validateForm()">
<input type="hidden" name="TabID" value="<%=oCategoryID%>">
<div align="center">
<table border="0">
<tr>
<td valign="top">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
          <td align="right">Name</td>
          <td>
  <input type="text" name="Name" size="40" maxlength="50" value="<%=Server.HTMLEncode(recString(oName))%>" class="required"></td>
        </tr>
    <tr>
          <td align="right">Sort Name</td>
          <td>
  <input type="text" name="SortName" size="40" maxlength="50" value="<%=Server.HTMLEncode(recString(oSortName))%>"></td>
        </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td>
  <textarea rows="7" name="ShortDesc" cols="35"><%=Server.HTMLEncode(recString(oShortDesc))%></textarea></td>
  </tr>
  <tr>
  <td align="right" valign="top">Picture</td>
  <td><input type="file" name="PictureFile" size="30">
<%

	DIM	sPictureName
	sPictureName = recString(oPicture)
	IF "" <> sPictureName THEN
		Response.Write "<br><input type=""checkbox"" name=""PictureDelete"" value=""ON"" id=""cbPictureDelete"">"
		Response.Write "<label for=""cbPictureDelete""> Delete Picture</label><br>" & vbCRLF
		Response.Write "<img src=""picture.asp?table=tabs&name=" & sPictureName & """>" & vbCRLF
	END IF

%>
  </td>
  </tr>
  <tr>
  <td>Page Sort
  </td>
  <td>
<%
DIM	oRS

sSelect = "" _
	&	"SELECT " _
	&		"RID, " _
	&		"ListName " _
	&	"FROM " _
	&		"pagelistoptions " _
	&	"ORDER BY " _
	&		"RID " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )

IF 0 < oRS.RecordCount THEN
	DIM	oID
	SET oID = oRS.Fields("RID")
	SET oName = oRS.Fields("ListName")
%>
<select name="PageSort">
<%
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oID & """"
		IF recNumber(oID) = recNumber(oTabSortID) THEN Response.Write " selected"
		Response.Write ">" & Server.HTMLEncode(oName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
%>
</select>
<%
END IF
oRS.Close

SET oRS = Nothing

%>
  </td>
  </tr>
  <tr>
  <td></td>
      <td>
  <input type="checkbox" name="Disabled" value="ON" <%=isTrue(oDisabled)%>>Disabled</td>
    </tr>
  </table>
</td>
<td valign="top">
<%


'DIM	sSelect

sSelect = "" _
	&	"SELECT " _
	&		"categories.CategoryID AS CategoryID, " _
	&		"categories.Name AS Name, " _
	&		"tabcategorymap.TabID AS TabID " _
	&	"FROM " _
	&		"categories " _
	&		"LEFT JOIN tabcategorymap " _
	&			"ON (tabcategorymap.CategoryID = categories.CategoryID " _
	&				"AND tabcategorymap.TabID = " & sTabID & ") " _
	&	"WHERE " _
	&		"0 = Disabled " _
	&	"ORDER BY " _
	&		"Name " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )

IF 0 < oRS.RecordCount THEN

	DIM	oTab
	'DIM	oName
	
	SET oID = oRS.Fields("CategoryID")
	SET oName = oRS.Fields("Name")
	SET oTab = oRS.Fields("TabID")

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
		Response.Write "<option value=""" & oID & """"
		IF NOT ISNULL(oTab) THEN Response.Write " selected"
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

%>
</td>
</tr>
</table>
  <p>
  <input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
  <input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
</div>

</form>

<%

		RS.Close

		IF FALSE THEN
		sSelect = "SELECT " _
			&		"COUNT(*) AS TotalMS " _
			&	"FROM Changes " _
			&	"WHERE " _
			&		"ProductGroup=" & ProductGroupNo & " " _
			&		"AND Deliverable=" & Component & " " _
			&		"AND CurrentStatus<400 " _
			&	""


		With RS
			.CursorLocation = 3 'adUseClient
			.CursorType     = 3 'adOpenStatic
			.LockType       = 3 'adLockOptimistic
			.Open sSelect, g_sDSN
			Set .ActiveConnection = Nothing
		End With
		IF NOT RS.EOF Then
			DIM	nTotal
			
			nTotal = CLNG(RS("TotalMS"))
			Response.Write "<p><a href=""results.asp?ProductGroup=" & ProductGroupNo _
				&	"&Component=" & Component _
				&	"&StatusOp=LT&Status=400&voEdit=ON"" target=""_blank"">" & nTotal & " Total (Active)</a></p>"
		END IF
		RS.Close
		END IF 'false

	END IF
	
SET	RS = Nothing
	
%>

<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->