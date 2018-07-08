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
<!--#include file="scripts\include_picture.asp"-->
<%


DIM	sID
DIM	sCategory
DIM	sBrand
sID = Request("page")
sCategory = Request("category")




%><html>

<head>
<title>Image Library</title>
<meta name="navigate" content="tab">
<meta name="sortname" content="zzzzzz">
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="scripts/formvalidate.js"></script>
<script type="text/javascript" language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "./";
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
		var re = /^[\w\.-]+$/gi;
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

<body bgcolor="#FFFFFF" id="pageEdit">
<!--#include file="scripts/index_include.asp"-->

<div style="text-align: right">The Image library identifies images or pictures that can be shared between pages.&nbsp; To 
	specify a picture from the library, prefix the label with a tilde (&quot;~&quot;).</div>

<h1>Image Library</h1>
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


	DIM	nPageID
	nPageID = 0


	sSelect = "" _
	&	"SELECT RID " _
	&	"FROM pages " _
	&		"INNER JOIN categories ON categories.CategoryID = pages.Category " _ 
	&	"WHERE " _
	&		"categories.Name='@Library' " _
	&		"AND pages.PageName='~Library' " _
	&	";"
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		nPageID = recNumber( oRS.Fields("RID") )
		oRS.Close
		SET oRS = Nothing
	END IF






%>
<form method="POST" action="imagessubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
	<input type="hidden" name="RID" value="<%=nPageID %>">
	<input type="hidden" name="QualCategory" value="<%=sCategory%>">
	<div align="center">
		<table border="0" cellpadding="2">
			<tr>
				<td bgcolor="#CCCCCC"><b>Picture Label</b><br>
				(required if uploading)<br>
				must use only:<br>
				&nbsp;letters, numbers, dash, <br>
				period and underscore</td>
				<td valign="top" bgcolor="#CCCCCC"><b>Specify Picture File</b><br>
				Use &quot;Browse...&quot; to <br>
				locate pictures on your computer</td>
		<td valign="top" bgcolor="#CCCCCC" colspan="2">Maximum<br>width / Height<br>
		pictures scaled<br>to fit (max: 640)</td>
			</tr>
			<%
	DIM	sSelect
	DIM	oRS
	DIM	n
	FOR n = 1 TO 5
%>
			<tr>
				<td>
				<input type="text" name="Label<%=n%>" size="20" onchange="verifyLabel(this)">
				</td>
				<td><input type="file" name="LabelPictureFile<%=n%>" size="30"> </td>
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
			&		"ImageID " _
			&	"FROM " _
			&		"(pictures " _
			&		"INNER JOIN pages ON pages.RID = pictures.PageID) " _
			&		"INNER JOIN categories ON categories.CategoryID = pages.Category " _ 
			&	"WHERE " _
			&		"categories.Name='@Library' " _
			&		"AND pages.PageName='~Library' " _
			&		"AND pictures.PageID = pages.RID " _
			&	"ORDER BY " _
			&		"Label" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oPictureID
	DIM	oPageID
	DIM	oLabel
	DIM	oImageID

	SET oPictureID = oRS.Fields("PictureID")
	SET oPageID = oRS.Fields("PageID")
	SET oLabel = oRS.Fields("Label")
	SET oImageID = oRS.Fields("ImageID")
	
	IF 0 < oRS.RecordCount THEN
%>
		<table cellpadding="2">
			<tr>
				<th bgcolor="#CCCCCC">Check to<br>
				Delete</th>
				<th bgcolor="#CCCCCC">Existing Pictures</th>
			</tr>
			<%
		DIM nImageID
		n = 0
		oRS.MoveFirst
		DO UNTIL oRS.EOF
			nImageID = recNumber(oImageID)
			IF 0 < nImageID THEN
				n = n + 1
				Response.Write "<tr>"
				Response.Write "<td valign=""top"" align=""center"" bgcolor=""#EEEEEE"">"
				Response.Write "<input type=""checkbox"" name=""LabelPictureDelete" & n & """ value=""" & oPictureID & """ id=""cbLabelPictureDelete" & n & """>"
				Response.Write "</td>" & vbCRLF
				Response.Write "<th align=""left"">"
				Response.Write "~" & Server.HTMLEncode(recString(oLabel)) & "<br>" & vbCRLF
				'Response.Write sPictureName & "<br>" & vbCRLF
				Response.Write "<img src=""../picture.asp?id=" & nImageID & """>"
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
	
  %>
		<p><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>
<%



%> 

<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
