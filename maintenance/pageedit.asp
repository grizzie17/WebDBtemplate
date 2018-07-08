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
<!--#include file="scripts\include_picture.asp"-->
<%


DIM	sID
DIM	sCategory
DIM	sBrand
sID = Request("page")
sCategory = Request("category")




%><html>

<head>
<title>Edit Page</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="scripts/formvalidate.js"></script>
<script language="JavaScript" type="text/javascript" src="scripts/pageedit.js"></script>
<script type="text/javascript" language="JavaScript">
<!--

	function doCancel() {
		window.location.href = "pages.asp?category=<%=sCategory%>";
	}

	function validateForm() {
		var oForm = document.EditForm;
		var sts;

		sts = validateRequired(oForm);
		if (sts)
			sts = checkRequiredPictureLabel(oForm);
		return sts;
	}


	function checkRequiredPictureLabel(theForm) {
		var bMissing = false;
		var i;
		var oField;
		var sValue;

		i = 0;
		while (i++ < 5) {
			oField = eval(theForm.name + ".LabelPictureFile" + i);
			sValue = oField.value;
			oField = eval(theForm.name + ".Label" + i);
			if (!verifyLabel(oField))
				bMissing = true;
			if (0 < sValue.length) {
				if (oField.value.length < 1)
					bMissing = true;
			}
			else {
				if (0 < oField.value.length)
					bMissing = true;
			}
		}

		if (bMissing) {
			alert("Uploading Pictures requires\nboth a label and the picture.");

			// false causes the form submission to be canceled
			return false;
		}
		else {
			return true;
		}
	}




	function verifyLabel(oLabel) {
		var sValue = oLabel.value;
		if (0 < sValue.length) {
			var re = /^[\w\.-]+$/gi;
			var ar;
			ar = sValue.match(re);
			if (null == ar) {
				alert("Picture label contains invalid characters.");
				oLabel.focus();
				return false;
			}
			else {
				return true;
			}
		}
		else {
			return true;
		}
	}


	function onSelectFormat(oThis) {
		var i = oThis.selectedIndex;
		var s = oThis.options[i].value;
		if (s == "HTML") {
			var oTextarea = document.getElementById("Body");
			if (0 == oTextarea.innerHTML.length)
				oTextarea.innerHTML = "<html><body></body></html>";
			CKEDITOR.replace(oTextarea, {
				extraPlugins: 'uploadimage',
				uploadUrl: 'pageupload.asp?type=files&amp;responsetype=json',
				filebrowserBrowseUrl: 'pageimagebrowse.asp',
				filebrowserImageBrowseUrl: 'pageimagebrowse.asp?type=images',
				filebrowserUploadUrl: 'pageupload.asp?type=files',
				filebrowserImageUploadUrl: 'pageupload.asp?type=images',
				fullPage: true,
				allowedContent: true
			});
		}
		else {
			if (CKEDITOR) {
				if (CKEDITOR.instances) {
					var o = CKEDITOR.instances['Body'];
					if (o)
						o.destroy();
				}
			}
		}

		return true;
	}


	function onLoadTextareaBody() {
		var oSelectFormat = document.getElementById("Format");
		return onSelectFormat(oSelectFormat);
	}
	if (window.addEventListener)
		window.addEventListener("load", onLoadTextareaBody, false);
	else if (window.attachEvent)
		window.attachEvent("onload", onLoadTextareaBody);
	else
		setTimeout("onLoadTextareaBody()", 1.0 * 1000);


	//-->
</script>
<script type="text/javascript" src="ckeditor/ckeditor.js"></script>
</head>

<body bgcolor="#FFFFFF" id="pageEdit">

<h1>Edit Page</h1>
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

FUNCTION makeSelect( x, y )
	IF UCASE(x) = UCASE(y) THEN
		makeSelect = " selected "
	ELSE
		makeSelect = " "
	END IF
END FUNCTION


DIM	g_sMimeList
DIM g_sMimeGraphic
DIM g_sMimeOther
g_sMimeList = ""
g_sMimeGraphic = ""
g_sMimeOther = ""

SUB BuildMimeLists()

	DIM	sSuffix
	DIM	aSuffix
	DIM	aMime
	DIM	i
	DIM	oFile
	DIM	sFile
	DIM	sData
	DIM	nSwitch
	nSwitch = 0
	
	sFile = Server.MapPath("scripts/mime.txt")
	IF g_oFSO.FileExists(sFile) THEN
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT oFile IS Nothing THEN
			sData = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			aMime = SPLIT(sData, vbCRLF )
			DIM	sLine
			FOR EACH sLine IN aMime
				IF "" <> sLine THEN
					SELECT CASE nSwitch
					CASE 0
						IF "-" = sLine THEN
							nSwitch = 1
						ELSE
							aSuffix = SPLIT(sLine,vbTAB)
							IF LBOUND(aSuffix) < UBOUND(aSuffix) THEN
								g_sMimeList = g_sMimeList & "." & aSuffix(0)
								g_sMimeGraphic = g_sMimeGraphic & "." & aSuffix(0)
							END IF
						END IF
					CASE 1
						aSuffix = SPLIT(sLine,vbTAB)
						IF LBOUND(aSuffix) < UBOUND(aSuffix) THEN
							g_sMimeList = g_sMimeList & "." & aSuffix(0)
							g_sMimeOther = g_sMimeOther & "." & aSuffix(0)
						END IF
					CASE ELSE
					END SELECT
				END IF
			NEXT
			g_sMimeList = g_sMimeList & "."
			g_sMimeGraphic = g_sMimeGraphic & "."
			g_sMimeOther = g_sMimeOther & "."
		END IF
	END IF

END SUB
BuildMimeLists





	DIM	sSelect
	sSelect = "" _
			&	"SELECT " _
			&		"pages.*, " _
			&		"schedules.DateBegin, " _
			&		"schedules.DateEnd, " _
			&		"schedules.DateEvent " _
			&	"FROM " _
			&		"(pages " _
			&		"LEFT JOIN schedules ON pages.RID = schedules.PageID) " _
			&	"WHERE " _
			&		"pages.RID=" & sID & " " _
			&	"ORDER BY " _
			&		"PageName " _
			&	";"
	
	DIM	oRS
	DIM	RS
	SET RS = dbQueryRead( g_DC, sSelect )



	IF RS.EOF Then
		RS.Close
		Response.Write "<p>Requested Page not found</p>"
	ELSE
		DIM	oRID

		DIM	oID
		DIM	oPageName
		DIM	oSortName
		DIM	oShortDesc
		DIM	oFormat
		DIM	oTitle
		DIM	oDesc
		DIM	oBody
		DIM	oCategory
		DIM	oDateCreated
		DIM	oDateModified
		DIM	oPicture
		DIM	oDisabled
		
		DIM	oDateBegin
		DIM oDateEnd
		DIM	oDateEvent
		
		SET oRID = RS.Fields("RID")
		SET	oPageName = RS.Fields("PageName")
		SET oSortName = RS.Fields("SortName")
		SET oFormat = RS.Fields("Format")
		SET	oTitle = RS.Fields("Title")
		SET oDesc = RS.Fields("Description")
		SET oBody = RS.Fields("Body")
		SET oCategory = RS.Fields("Category")
		SET oDateCreated = RS.Fields("DateCreated")
		SET oDateModified = RS.Fields("DateModified")
		SET oPicture = RS.Fields("Picture")
		SET	oDisabled = RS.Fields("Disabled")
		
		SET oDateBegin = RS.Fields("DateBegin")
		SET oDateEnd = RS.Fields("DateEnd")
		SET oDateEvent = RS.Fields("DateEvent")
		
		DIM	sFormat
		sFormat = recString(oFormat)
		IF "" = sFormat THEN
			sFormat = "STFT"
		END iF


%>

<form method="POST" action="pageeditsubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
<input type="hidden" name="RID" value="<%=recNumber(oRID)%>">
<input type="hidden" name="PictureSuffixes" value="<%=g_sMimeList%>">
<input type="hidden" name="CategoryID" value="<%=recNumber(oCategory)%>">
<input type="hidden" name="QualCategory" value="<%=sCategory%>">
<div align="center">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
  <td align="right">Page</td>
      <td>
  <input type="checkbox" name="Disabled" value="ON" <%=checkTrue(oDisabled)%> id="DisabledCheck"><label for="DisabledCheck">Disabled</label></td>
    </tr>
    <tr>
  <td align="right" valign="top">Category</td>
  <td><select size="1" name="Category" class="required">
  <%
	sSelect = "SELECT " _
			&		"CategoryID, " _
			&		"Name " _
			&	"FROM " _
			&		"categories " _
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
		Response.Write ">" & Server.HTMLEncode(recString(oCatName)) & "</option>" & vbCRLF
		oRS.MoveNext
	LOOP
	
	oRS.Close
	SET oRS = Nothing
  %>
  </select></td>
    </tr>
    <tr>
          <td align="right">Name</td>
          <td>
	  <input type="text" name="PageName" size="40" maxlength="<%=oPageName.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oPageName))%>" class="required"></td>
    </tr>
    <tr>
          <td align="right">Sort Name</td>
          <td>
  <input type="text" name="SortName" size="20" maxlength="<%=oSortName.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oSortName))%>"></td>
        </tr>
    <tr>
          <td align="right">Format</td>
          <td><select name="Format" id="Format" class="required" onload="onSelectFormat(this);" onchange="onSelectFormat(this);">
			<option <%=makeSelect(sFormat,"STFT")%> value="STFT">Simple Text Format</option>
			<option <%=makeSelect(sFormat,"4MTX")%> value="STFT">Forum Text Format</option>
			<option <%=makeSelect(sFormat,"HTML")%> value="HTML">Hyper-Text</option>
			<option <%=makeSelect(sFormat,"TEXT")%> value="TEXT">Plain-Text</option>
			</select>&nbsp;&nbsp; <a target="_blank" href="scripts/htmlformathelp.asp">Simple Text Format Help</a></td>
        </tr>
    <tr>
          <td align="right">Title</td>
          <td>
		  <input type="text" name="Title" id="Title" size="55" maxlength="<%=oTitle.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oTitle))%>" class="required">
			<a href="javascript:fixupTextInput('Title')">fix-special-chars</a></td>
    </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td>
  <input type="text" name="Description" id="Description" size="55" maxlength="<%=oDesc.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oDesc))%>"> <a href="javascript:fixupTextInput('Description')">
	fix-special-chars</a></td>
  </tr>
  <tr>
  <td align="right" valign="top">Body</td>
  <td>
  <textarea rows="15" name="Body" id="Body" cols="80" wrap="soft"><%=Server.HTMLEncode(recString(oBody))%></textarea><br>
	<a href="javascript:fixupTextArea('Body')">fix-special-chars</a></td>
  </tr>
  <tr>
  <td align="right" valign="top">Date Range</td>
  <td>
  Begin:<input type="text" name="DateBegin" size="12" value="<%=Server.HTMLEncode(recString(oDateBegin))%>">
  End:<input type="text" name="DateEnd" size="12" value="<%=Server.HTMLEncode(recString(oDateEnd))%>">
  Event:<input type="text" name="DateEvent" size="12" value="<%=Server.HTMLEncode(recString(oDateEvent))%>">
  </td>
  </tr>
  <tr>
  <td align="right" valign="top">Picture</td>
  <td><input type="file" name="PictureFile" size="30">&nbsp; Max: w:<input type="text" name="PictureMaxW" size="4" value="300"> 
  h:<input type="text" name="PictureMaxH" size="4" value="400">
<%

	DIM	nPictureID
	nPictureID = recNumber(oPicture)
	IF 0 < nPictureID THEN
		Response.Write "<br><input type=""checkbox"" name=""PictureDelete"" value=""ON"" id=""cbPictureDelete"">"
		Response.Write "<label for=""cbPictureDelete""> Delete Picture</label><br>" & vbCRLF
		Response.Write "<b>default</b><br>" & vbCRLF
		Response.Write "<img src=""../picture.asp?id=" & nPictureID & """>" & vbCRLF
	END IF

%>
  </td>
  </tr>
  <tr>
  <td valign="top">Labeled Pictures<br />
      <a target="_blank" href="scripts/htmlpicturehelp.asp">Picture Help</a>
  </td>
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
		<td valign="top" bgcolor="#CCCCCC" colspan="2">Maximum<br>width / Height<br>
		pictures scaled<br>to fit (max: 640)</td>
    </tr>
<%
	DIM	n
	FOR n = 1 TO 5
%>
    <tr>
      <td><input type="text" name="Label<%=n%>" size="20" onblur="verifyLabel(this)" onchange="verifyLabel(this)">
      </td>
      <td><input type="file" name="LabelPictureFile<%=n%>" size="30">
      </td>
      <td>w:<input type="text" name="labelPictureMaxW<%=n%>" size="4" value="480"></td>
      <td>h:<input type="text" name="labelPictureMaxH<%=n%>" size="4" value="480"></td>
    </tr>
  
<%
	NEXT 'n
%>
  </table>
<%
	sSelect = "SELECT " _
			&		"PictureID, " _
			&		"PageID, " _
			&		"Label, " _
			&		"ImageID, " _
			&		dbQ("File") & " " _
			&	"FROM " _
			&		"(pictures " _
			&		"LEFT JOIN images ON pictures.ImageID = images.RID) " _
			&	"WHERE " _
			&		"PageID = " & recNumber(oRID) & " "  _
			&	"ORDER BY " _
			&		"Label" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	sPictureName
	DIM	oPictureID
	DIM	oPageID
	DIM	oLabel
	DIM	oImageID
	DIM	oFile

	SET oPictureID = oRS.Fields("PictureID")
	SET oPageID = oRS.Fields("PageID")
	SET oLabel = oRS.Fields("Label")
	SET oImageID = oRS.Fields("ImageID")
	SET oFile = oRS.Fields("File")
	
	IF 0 < oRS.RecordCount THEN
%>
		<table cellpadding="2">
			<tr>
			<th bgcolor="#CCCCCC">Check to<br>Delete</th>
			<th bgcolor="#CCCCCC">Existing Pictures</th>
			</tr>
<%
		DIM	i
		DIM	sSuffix
		n = 0
		oRS.MoveFirst
		DO UNTIL oRS.EOF
			sPictureName = recString(oFile)
			IF 0 < recNumber(oImageID) THEN
				n = n + 1
				i = INSTRREV(sPictureName,".")
				IF 0 < i THEN
					sSuffix = LCASE(MID(sPictureName,i+1))
				ELSE
					sSuffix = ""
				END IF
				Response.Write "<tr>"
				Response.Write "<td valign=""top"" align=""center"" bgcolor=""#EEEEEE"">"
				Response.Write "<input type=""checkbox"" name=""LabelPictureDelete" & n & """ value=""" & recNumber(oPictureID) & """ id=""cbLabelPictureDelete" & n & """>"
				Response.Write "</td>" & vbCRLF
				Response.Write "<th align=""left"">"
				Response.Write Server.HTMLEncode(recString(oLabel)) & "<br>" & vbCRLF
				IF 0 < INSTR( g_sMimeGraphic, "."&sSuffix&"." ) THEN
					Response.Write "<img src=""../picture.asp?id=" & recNumber(oImageID) & """>"
				ELSEIF 0 < INSTR( g_sMimeOther, "."&sSuffix&"." ) THEN
					Response.Write "<img src=""images/file_" & sSuffix & ".gif"">"
				ELSE
					Response.Write "<img src=""images/file_Unknown.gif"">"
				END IF
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
    <tr>
          <td align="right">Dates</td>
          <td>Created: <%=recString(oDateCreated)%>, Last Modified: <%=recString(oDateModified)%>
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


<!--#include file="scripts\byline.asp"-->

</body>

</html>
<%
SET	RS = Nothing
%><!--#include file="scripts\include_db_end.asp"-->