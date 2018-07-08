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
<!--#include file="scripts\include_picture.asp"--><%



IF NOT authHasAccess( "SRVe" ) THEN
	Response.Redirect "./"
END IF




DIM	sID
DIM	sCategory
DIM	sBrand
sID = Request("page")
sCategory = Request("category")




%><html>

<head>
<title>Edit Server Configuration</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="scripts/formvalidate.js"></script>
<script language="JavaScript" type="text/javascript" src="scripts/pageedit.js"></script>
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
	//if ( sts )
	//	sts = checkRequiredPictureLabel( oForm );
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

<body bgcolor="#FFFFFF" id="pageEdit">

<h1>Edit Server Configuration</h1>
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





	DIM	sSelect
	sSelect = "" _
			&	"SELECT " _
			&		"* " _
			&	"FROM Config " _
			&	";"
	
	DIM	oRS
	DIM	RS
	SET RS = dbQueryRead( g_DC, sSelect )



	IF RS.EOF Then
		RS.Close
		Response.Write "<p>Requested Config Data not found</p>"
	ELSE
		DIM	oRID

		DIM	oID
		DIM	oDomain
		DIM	oSiteName
		DIM	oSiteTitle
		DIM	oSiteOrgDesignation
		DIM oSiteChapter
		DIM	oSiteChapterID
		DIM	oSiteMotto
		DIM	oSiteLocation
		DIM	oSiteZip
		DIM	oCopyrightStartYear
		DIM	oMailServer
		DIM	oMailUser
		DIM	oMailPW
		DIM	oSiteTabAnnouncements
		DIM	oSiteTabNewsletters
		DIM	oPicture
		DIM	oDisabled
		
		SET oRID = RS.Fields("RID")
		SET	oDomain = RS.Fields("Domain")
		SET oSiteName = RS.Fields("SiteName")
		SET oSiteTitle = RS.Fields("SiteTitle")
		SET	oSiteOrgDesignation = RS.Fields("SiteOrgDesignation")
		SET oSiteChapter = RS.Fields("SiteChapter")
		SET oSiteChapterID = RS.Fields("SiteChapterID")
		SET	oSiteMotto = RS.Fields("SiteMotto")
		SET oSiteLocation = RS.Fields("SiteLocation")
		SET oSiteZip = RS.Fields("SiteZip")
		SET oCopyrightStartYear = RS.Fields("CopyrightStartYear")
		SET oMailServer = RS.Fields("MailServer")
		SET oMailUser = RS.Fields("MailUser")
		SET oMailPW = RS.Fields("MailPW")
		SET oSiteTabAnnouncements = RS.Fields("SiteTabAnnouncements")
		SET oSiteTabNewsletters = RS.Fields("SiteTabNewsletters")
		'SET oPicture = RS.Fields("Picture")
		'SET	oDisabled = RS.Fields("Disabled")
		


%>
<form method="POST" action="configsrveditsubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
	<input type="hidden" name="RID" value="<%=oRID%>">
	<div align="center">
		<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
			<tr>
				<td align="right">Domain</td>
				<td>
				<input type="text" name="Domain" size="40" maxlength="<%=oDomain.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oDomain))%>" class="required"></td>
			</tr>
			<tr>
				<td align="right">Mail Server</td>
				<td>
				<input type="text" name="MailServer" size="40" maxlength="<%=oMailServer.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oMailServer))%>"></td>
			</tr>
			<tr>
				<td align="right">Mail User</td>
				<td>
				<input type="text" name="MailUser" size="40" maxlength="<%=oMailUser.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oMailUser))%>"></td>
			</tr>
			<tr>
				<td align="right">Mail Password</td>
				<td>
				<input type="password" name="MailPW" size="40" maxlength="<%=oMailPW.DefinedSize%>" value="<%=Server.HTMLEncode(recString(oMailPW))%>"></td>
			</tr>
<%
IF FALSE THEN ' disable pictures
%>
			<tr>
				<td align="right" valign="top">Picture</td>
				<td><input type="file" name="PictureFile" size="30"> <%

	DIM	sPictureName
	sPictureName = recString(oPicture)
	IF "" <> sPictureName THEN
		Response.Write "<br><input type=""checkbox"" name=""PictureDelete"" value=""ON"" id=""cbPictureDelete"">"
		Response.Write "<label for=""cbPictureDelete""> Delete Picture</label><br>" & vbCRLF
		Response.Write "<img src=""picture.asp?table=pages&name=" & sPictureName & """>" & vbCRLF
	END IF

%> </td>
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
						<td><input type="text" name="Label<%=n%>" size="20" onchange="verifyLabel(this)"> </td>
						<td><input type="file" name="LabelPictureFile<%=n%>" size="30"> </td>
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
			&		"File " _
			&	"FROM " _
			&		"Pictures " _
			&	"WHERE " _
			&		"PageID = " & oRID & " "  _
			&	"ORDER BY " _
			&		"Label" _
			&	";"
	
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	DIM	oPictureID
	DIM	oPageID
	DIM	oLabel
	DIM	oFile

	SET oPictureID = oRS.Fields("PictureID")
	SET oPageID = oRS.Fields("PageID")
	SET oLabel = oRS.Fields("Label")
	SET oFile = oRS.Fields("File")
	
	IF 0 < oRS.RecordCount THEN
%>
				<table cellpadding="2">
					<tr>
						<th bgcolor="#CCCCCC">Check to<br>
						Delete</th>
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
				Response.Write "<img src=""picture.asp?table=pages&name=" & sPictureName & """>"
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
  %> </td>
			</tr>
<%
END IF	' disable pictures
%>
		</table>
		<p><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>
<%

		RS.Close


	END IF
%> <!--webbot bot="Include" u-include="../_private/byline.htm" tag="BODY" startspan -->

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

<!--#include file="scripts\byline.asp"-->

</body>

</html>
<%
SET	RS = Nothing
%><!--#include file="scripts\include_db_end.asp"-->