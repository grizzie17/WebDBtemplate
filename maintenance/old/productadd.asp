<%@ LANGUAGE="VBSCRIPT" %> <%
Option Explicit
Response.Expires = 0
Response.Buffer = True



%>
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
					&		"Name " _
					&	"FROM " _
					&		"Categories " _
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
	sCategoryName = "Undefined"
END IF

%><html>

<head>
<title>Add Page (<%=sCategoryName%>)</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="../theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="../scripts/formvalidate.js"></script>
<script language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "products.asp?category=<%=sCategory%>";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}


//-->
</script>
</head>

<body bgcolor="#FFFFFF">

<h1>Add Page (<%=sCategoryName%>)</h1>
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






%>
<form method="POST" action="productaddsubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
	<input type="hidden" name="QualCategory" value="<%=sCategory%>">
	<input type="hidden" name="QualBrand" value="<%=sBrand%>">
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
			&		"Categories " _
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
	
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oCategoryID & """"
		IF "" <> sCategory THEN
			IF recNumber(oCategoryID) = CLNG(sCategory) THEN
				Response.Write " selected"
				IF 0 <> oDisabled THEN
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
				<td align="right">Page-ID</td>
				<td><input type="text" name="ProdID" size="20" maxlength="32" value="" class="required"></td>
			</tr>
			<tr>
				<td align="right">Name</td>
				<td><input type="text" name="Name" size="40" maxlength="40" value="" class="required"></td>
			</tr>
			<tr>
				<td align="right">Sort Name</td>
				<td><input type="text" name="SortName" size="20" maxlength="20" value=""></td>
			</tr>
			<tr>
				<td align="right">Title</td>
				<td><input type="text" name="ShortDesc" size="50" maxlength="80" value="" class="required"></td>
			</tr>
			<tr>
				<td align="right">Format</td>
				<td><select name="Format">
				<option value="TXHT">Simple Text Format</option>
				<option value="HTML">Hyper-Text</option>
				<option value="TEXT">Plain-Text</option>
				</select> </td>
			</tr>
			<tr>
				<td align="right" valign="top">Description</td>
				<td><textarea rows="10" name="Desc" cols="50"></textarea></td>
			</tr>
			<tr>
				<td align="right" valign="top">Brand</td>
				<td><select size="1" name="Brand"><%
sWhereCat = ""
IF "" <> sBrand THEN
	sWhereCat = "OR BrandID=" & sBrand & " "
END IF
	sSelect = "SELECT " _
			&		"BrandID, " _
			&		"Name, " _
			&		"Disabled " _
			&	"FROM " _
			&		"Brands " _
			&	"WHERE " _
			&		"0 = Disabled " _
			&		sWhereCat _
			&	"ORDER BY " _
			&		"Name" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oBrandID
	DIM	oName
	DIM	oDisabled
	
	SET oBrandID = oRS.Fields("BrandID")
	SET oName = oRS.Fields("Name")
	SET oDisabled = oRS.Fields("Disabled")
	
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oBrandID & """"
		IF "" <> sBrand THEN
			IF recNumber(oBrandID) = CLNG(sBrand) THEN
				Response.Write " selected"
				IF 0 <> oDisabled THEN
					Response.Write " style=""color: #CCCCCC"""
				END IF
			END IF
		END IF
		Response.Write ">" & Server.HTMLEncode(oName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
	
	oRS.Close
	SET oRS = Nothing
  %></select></td>
			</tr>
			<%
IF FALSE THEN
%>
			<tr>
				<td align="right" valign="top">Unit-Price</td>
				<td>
				<table border="0" cellpadding="2" id="table1" cellspacing="0">
					<tr>
						<td align="center">
						<input type="text" name="UnitPrice" size="8" maxlength="50" value="" class="required"></td>
						<td align="center"><input type="text" name="ListPrice" size="8" maxlength="50" value=""></td>
						<td align="center"><input type="text" name="MAPrice" size="8" maxlength="50" value=""></td>
						<td align="center"><input type="text" name="Cost" size="8" maxlength="50" value=""></td>
					</tr>
					<tr>
						<td align="center">Unit</td>
						<td align="center">List</td>
						<td align="center">M.A.P.</td>
						<td align="center">Cost</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td align="right">Weight</td>
				<td><input type="text" name="Weight" size="8" maxlength="50" value=""></td>
			</tr>
			<tr>
				<td align="right">Quantity</td>
				<td><input type="checkbox" name="HideQuantity" id="HideQuantityCheck" value="ON"><label for="HideQuantityCheck"> 
				Hide</label></td>
			</tr>
			<%
END IF
%>
			<tr>
				<td align="right" valign="top">Website</td>
				<td><input type="text" name="Website" size="40"></td>
			</tr>
			<tr>
				<td align="right" valign="top">Picture</td>
				<td><input type="file" name="PictureFile" size="30"></td>
			</tr>
			<tr>
				<td align="right" valign="top">Large Picture</td>
				<td><input type="file" name="LargePictureFile" size="30"></td>
			</tr>
			<tr>
				<td align="right" valign="top">Thumbnail</td>
				<td><input type="file" name="ThumbnailFile" size="30"></td>
			</tr>
			<tr>
				<td align="right">Date</td>
				<td><input type="text" name="DateAdded" size="10" maxlength="12" value=""></td>
			</tr>
		</table>
		<p><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>
<!--webbot bot="Include" u-include="../../_private/byline.htm" tag="BODY" startspan -->

<script language="JavaScript">
<!--

function makeByLine()
{
	document.write( '<' + 'script language="JavaScript" src="http://' );
	if ( "localhost" == location.hostname )
	{
		document.writeln( 'localhost/BearConsultingGroup' );
	}
	else
	{
		document.writeln( 'BearConsultingGroup.com' );
	}
	document.writeln( '/designby_small.js"></' + 'script>' );
}
makeByLine()

//-->
</script>
<!--script language="JavaScript" src="http://BearConsultingGroup.com/designbyadvert.js"></script-->

<!--webbot bot="Include" i-checksum="20551" endspan -->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->