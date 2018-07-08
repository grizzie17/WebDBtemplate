<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<%


DIM	sID
DIM	sCategory
DIM	sBrand
sID = Request("product")
sCategory = Request("category")
sBrand = Request("brand")




%><html>

<head>
<title>Edit Product</title>
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
	var oForm = document.EditForm;
	var sts;
	
	sts = validateRequired( oForm );
	if ( sts )
		sts = checkRequiredPictureLabel( oForm );
	return sts;
}


function checkRequiredPictureLabel( theForm )
{
    var bMissing = false;
    var i;
    var oField;
    var sValue;
    
    i = 0;
    while ( i++ < 5 )
    {
    	oField = eval( theForm.name+".LabelPictureFile" + i );
       	sValue = oField.value;
    	oField = eval( theForm.name+".Label" + i );
    	if ( ! verifyLabel( oField ) )
    		bMissing = true;
    	if ( 0 < sValue.length )
    	{
    		if ( oField.value.length < 1 )
    			bMissing = true;
    	}
    	else
    	{
    		if ( 0 < oField.value.length )
    			bMissing = true;
    	}
    }
            
    if ( bMissing )
    {
        alert( "Uploading Pictures requires\nboth a label and the picture.");
        
        // false causes the form submission to be canceled
        return false;
    }
    else
    {
        return true;
    }
}




function verifyLabel( oLabel )
{
	var sValue = oLabel.value;
	if ( 0 < sValue.length )
	{
		var re = /^[\w-]+$/gi;
		var ar;
		ar = sValue.match( re );
		if ( null == ar )
		{
			alert( "Picture label contains invalid characters.");
			oLabel.focus();
			return false;
		}
		else
		{
			return true;
		}
	}
	else
	{
		return true;
	}
}



//-->
</script>
</head>

<body bgcolor="#FFFFFF">

<h1>Edit Page</h1>
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
	sSelect = "" _
			&	"SELECT " _
			&		"Products.*, " _
			&		"Gallery.Body " _
			&	"FROM (Products " _
			&		"LEFT JOIN Gallery ON Gallery.ProductID = Products.RID) " _
			&	"WHERE " _
			&		"RID=" & sID & " " _
			&	"ORDER BY " _
			&		"Name " _
			&	";"
	
	DIM	oRS
	DIM	RS
	SET RS = dbQueryRead( g_DC, sSelect )



	IF RS.EOF Then
		RS.Close
		Response.Write "<p>Requested Brand not found</p>"
	ELSE
		DIM	oRID

		DIM	oID
		DIM	oName
		DIM	oSortName
		DIM	oShortDesc
		DIM	oFormat
		DIM	oDesc
		DIM	oCategory
		DIM	oBrand
		DIM	oUnitPrice
		DIM	oListPrice
		DIM	oMAPrice
		DIM	oCost
		DIM	oWeight
		DIM	oDateAdded
		DIM	oHideQuantity
		DIM	oOptions
		DIM	oStockCount
		DIM	oWebsite
		DIM	oPicture
		DIM	oLargePicture
		DIM	oThumbnail
		DIM	oDisabled
		DIM	oBody
		
		SET oRID = RS.Fields("RID")
		SET oID = RS.Fields("ProdID")
		SET	oName = RS.Fields("Name")
		SET oSortName = RS.Fields("SortName")
		SET oShortDesc = RS.Fields("ShortDesc")
		SET oFormat = RS.Fields("Format")
		SET oDesc = RS.Fields("Description")
		SET oCategory = RS.Fields("Category")
		SET oBrand = RS.Fields("Brand")
		SET	oUnitPrice = RS.Fields("UnitPrice")
		SET oListPrice = RS.Fields("ListPrice")
		SET oMAPrice = RS.Fields("MAPrice")
		SET oCost = RS.Fields("Cost")
		SET	oWeight = RS.Fields("Weight")
		SET oDateAdded = RS.Fields("DateAdded")
		SET oHideQuantity = RS.Fields("HideQuantity")
		SET oOptions = RS.Fields("Options")
		SET oStockCount = RS.Fields("StockCount")
		SET oWebsite = RS.Fields("URL")
		SET oPicture = RS.Fields("Picture")
		SET oLargePicture = RS.Fields("LargePicture")
		SET oThumbnail = RS.Fields("Thumbnail")
		SET	oDisabled = RS.Fields("Disabled")
		SET oBody = RS.Fields("Body")


%>

<form method="POST" action="producteditsubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
<input type="hidden" name="RID" value="<%=oRID%>">
<input type="hidden" name="CategoryID" value="<%=oCategory%>">
<input type="hidden" name="QualCategory" value="<%=sCategory%>">
<input type="hidden" name="QualBrand" value="<%=sBrand%>">
<div align="center">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
  <td align="right">Page</td>
      <td>
  <input type="checkbox" name="Disabled" value="ON" <%=isTrue(oDisabled)%> id="DisabledCheck"><label for="DisabledCheck">Disabled</label></td>
    </tr>
    <tr>
  <td align="right" valign="top">Category</td>
  <td><select size="1" name="Category" class="required">
  <%
	sSelect = "SELECT " _
			&		"CategoryID, " _
			&		"Name " _
			&	"FROM " _
			&		"Categories " _
			&	"WHERE " _
			&		"0 = Disabled " _
			&		"OR CategoryID=" & oCategory & " " _
			&	"ORDER BY " _
			&		"Name" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oCategoryID
	DIM oCatName
	
	SET oCategoryID = oRS.Fields("CategoryID")
	SET oCatName = oRS.Fields("Name")
	
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oCategoryID & """"
		IF recNumber(oCategoryID) = recNumber(oCategory) THEN
			Response.Write " selected"
		END IF
		Response.Write ">" & Server.HTMLEncode(oCatName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
	
	oRS.Close
	SET oRS = Nothing
  %>
  </select></td>
    </tr>
    <tr>
          <td align="right">Page-ID</td>
       <td>
		  <input type="text" name="ProdID" size="20" maxlength="50" value="<%=Server.HTMLEncode(recString(oID))%>" class="required"> <font color="#CCCCCC" size="-2">(<%=oRID%>)</font></td>
    </tr>
    <tr>
          <td align="right">Name</td>
          <td>
	  <input type="text" name="Name" size="40" maxlength="50" value="<%=Server.HTMLEncode(recString(oName))%>" class="required"></td>
    </tr>
    <tr>
          <td align="right">Sort Name</td>
          <td>
  <input type="text" name="SortName" size="20" maxlength="20" value="<%=Server.HTMLEncode(recString(oSortName))%>"></td>
        </tr>
    <tr>
          <td align="right">Title</td>
          <td>
		  <input type="text" name="ShortDesc" size="55" maxlength="80" value="<%=Server.HTMLEncode(recString(oShortDesc))%>" class="required"></td>
    </tr>
    <tr>
          <td align="right">Format</td>
          <td><select name="Format">
			<option value="TXHT">Simple Text Format</option>
			<option value="HTML">Hyper-Text</option>
			<option value="TEXT">Plain-Text</option>
			</select>
  </td>
        </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td>
  <textarea rows="12" name="Desc" cols="60" wrap="soft"><%=Server.HTMLEncode(recString(oDesc))%></textarea></td>
  </tr>
    <tr>
  <td align="right" valign="top">Brand</td>
  <td><select size="1" name="Brand">
  <%
	sSelect = "SELECT " _
			&		"BrandID, " _
			&		"Name " _
			&	"FROM " _
			&		"Brands " _
			&	"WHERE " _
			&		"0 = Disabled " _
			&	"ORDER BY " _
			&		"Name" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oBrandID
	
	SET oBrandID = oRS.Fields("BrandID")
	SET oName = oRS.Fields("Name")
	
	DO UNTIL oRS.EOF
		Response.Write "<option value=""" & oBrandID & """"
		IF recNumber(oBrandID) = recNumber(oBrand) THEN
			Response.Write " selected"
		END IF
		Response.Write ">" & Server.HTMLEncode(oName) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
	
	oRS.Close
	SET oRS = Nothing
  %>
  </select></td>
    </tr>
<%
IF FALSE THEN
%>
    <tr>
          <td align="right" valign="top">Price</td>
          <td>
  <table border="0" id="table1" cellspacing="0">
	<tr>
		<td align="center">
  <input type="text" name="UnitPrice" size="8" value="<%=oUnitPrice%>" class="required"></td>
		<td align="center">
  <input type="text" name="ListPrice" size="8" value="<%=oListPrice%>"></td>
		<td align="center">
  <input type="text" name="MAPrice" size="8" value="<%=oMAPrice%>"></td>
		<td align="center">
  <input type="text" name="Cost" size="8" value="<%=oCost%>"></td>
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
          <td>
  <input type="text" name="Weight" size="8" value="<%=oWeight%>"></td>
        </tr>
    <tr>
          <td align="right">Quantity</td>
          <td>
  <input type="checkbox" name="HideQuantity" <%=isTrue(oHideQuantity)%> id="HideQuantityCheck" value="ON"><label for="HideQuantityCheck"> Hide</label></td>
        </tr>
    <tr>
          <td align="right">Stock Count</td>
          <td>
  <input type="text" name="StockCount" size="8" value="<%=oStockCount%>"></td>
        </tr>
<%
END IF
%>
    <tr>
          <td align="right">Date Added</td>
          <td>
  <input type="text" name="DateAdded" size="20" maxlength="23" value="<%=recString(oDateAdded)%>"></td>
        </tr>
 <%
IF FALSE THEN
 %>
  	<tr>
  <td align="right" valign="top">Options</td>
  <td>
  <textarea rows="9" name="Options" cols="60" wrap="off"><%=Server.HTMLEncode(recString(oOptions))%></textarea><br>
	<span style="font-size:x-small;font-family:sans-serif">checkbox :&nbsp; <i>option-name</i> : checked : <i>caption</i> : required<br>
	number : <i>option-name</i> : <i>default-value</i> : <i>caption</i> : 
	required<br>
	text : <i>option-name</i> : <i>default-value</i> : <i>caption</i> : required<br>
	select : <i>option-name</i> : <i>v1=name1;v2=name2</i> : <i>caption</i> : 
	required</span></td>
  	</tr>
 <%
END IF
 %>
  <tr>
  <td align="right" valign="top">Website</td>
  <td>
  <input type="text" name="Website" size="40" value="<%=recString(oWebsite)%>"></td>
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
		Response.Write "<img src=""picture.asp?table=products&name=" & sPictureName & """>" & vbCRLF
	END IF

%>
  </td>
  </tr>
  <tr>
  <td align="right" valign="top">Large Picture</td>
  <td><input type="file" name="LargePictureFile" size="30">
<%

	sPictureName = recString(oLargePicture)
	IF "" <> sPictureName THEN
		Response.Write "<br><input type=""checkbox"" name=""LargePictureDelete"" value=""ON"" id=""cbLargePictureDelete"">"
		Response.Write "<label for=""cbLargePictureDelete""> Delete Large Picture</label><br>" & vbCRLF
		Response.Write "<img src=""picture.asp?table=products&name=" & sPictureName & """>" & vbCRLF
	END IF

%>
  </td>
  </tr>
  <tr>
  <td align="right" valign="top">Thumbnail</td>
  <td><input type="file" name="ThumbnailFile" size="30">
<%

	sPictureName = recString(oThumbnail)
	IF "" <> sPictureName THEN
		Response.Write "<br><input type=""checkbox"" name=""ThumbnailDelete"" value=""ON"" id=""cbThumbnailDelete"">"
		Response.Write "<label for=""cbThumbnailDelete""> Delete Thumbnail</label><br>" & vbCRLF
		Response.Write "<img src=""picture.asp?table=products&name=" & sPictureName & """>" & vbCRLF
	END IF

%>
  </td>
  </tr>
  <tr>
  <td align="right" valign="top">Gallery</td>
  <td>
  <textarea rows="12" name="Body" cols="60" wrap="soft"><%=Server.HTMLEncode(recString(oBody))%></textarea></td>
  </tr>
  <tr>
  <td valign="top">Labeled Pictures</td>
      <td>

  <table border="0" cellpadding="2">
    <tr>
      <td bgcolor="#CCCCCC"><b>Picture Label</b><br>
        (required if uploading)<br>
        must use only:<br>
        &nbsp;letters, numbers, and dash</td>
      <td valign="top" bgcolor="#CCCCCC"><b>Specify Picture File</b><br>
        Use &quot;Browse...&quot; to <br>
		locate pictures on your computer</td>
    </tr>
<%
	DIM	n
	FOR n = 1 TO 5
%>
    <tr>
      <td><input type="text" name="Label<%=n%>" size="20" onchange="verifyLabel(this)">
      </td>
      <td><input type="file" name="LabelPictureFile<%=n%>" size="30">
      </td>
    </tr>
  
<%
	NEXT 'n
%>
  </table>
<%
	sSelect = "SELECT " _
			&		"PictureID, " _
			&		"ProductID, " _
			&		"Label, " _
			&		"File " _
			&	"FROM " _
			&		"Pictures " _
			&	"WHERE " _
			&		"ProductID = " & oRID & " "  _
			&	"ORDER BY " _
			&		"Label" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oPictureID
	DIM	oProductID
	DIM	oLabel
	DIM	oFile

	SET oPictureID = oRS.Fields("PictureID")
	SET oProductID = oRS.Fields("ProductID")
	SET oLabel = oRS.Fields("Label")
	SET oFile = oRS.Fields("File")
	
	IF 0 < oRS.RecordCount THEN
%>
		<table cellpadding="2">
			<tr>
			<th bgcolor="#CCCCCC">Check to<br>Delete</th>
			<th bgcolor="#CCCCCC">Existing Pictures</th>
			</tr>
<%
		n = 0
		DO UNTIL oRS.EOF
			sPictureName = recString(oFile)
			IF "" <> sPictureName THEN
				n = n + 1
				Response.Write "<tr>"
				Response.Write "<td valign=""top"" align=""center"" bgcolor=""#EEEEEE"">"
				Response.Write "<input type=""checkbox"" name=""LabelPictureDelete"" value=""" & oPictureID & """ id=""cbLabelPictureDelete" & n & """>"
				Response.Write "</td>" & vbCRLF
				Response.Write "<th align=""left"">"
				Response.Write Server.HTMLEncode(recString(oLabel)) & "<br>" & vbCRLF
				Response.Write "<img src=""picture.asp?table=products&name=" & sPictureName & """>"
				Response.Write "</th>" & vbCRLF
				Response.Write "</tr>" & vbCRLF
			END IF
			oRS.MoveNext
		LOOP
%>
		</table>
<%
		Response.Write "<input type=""hidden"" name=""LabelPictureDeleteCount"" value=""" & n & """>" & vbCRLF
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


	END IF
%>

<!--webbot bot="Include" U-Include="../../_private/byline.htm" TAG="BODY" startspan -->

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
<%
SET	RS = Nothing
%><!--#include file="scripts\include_db_end.asp"-->