<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<%


DIM	sID
sID = Request("brand")




%><html>

<head>
<title>Edit Brand</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="../theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="../scripts/formvalidate.js"></script>
<script language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "brands.asp";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}



//-->
</script>
</head>

<body bgcolor="#FFFFFF">

<h1>Edit Brand</h1>
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
			&	" * " _
			&	"FROM Brands " _
			&	"WHERE " _
			&		"BrandID=" & sID & " " _
			&	"ORDER BY " _
			&		"Name " _
			&	";"
	
	
	DIM	RS
	SET RS = dbQueryRead( g_DC, sSelect )



	IF RS.EOF Then
		RS.Close
		Response.Write "<p>Requested Brand not found</p>"
	ELSE
		DIM	oRID

		
		DIM	oID
		DIM	oName
		DIM	oDesc
		DIM	oWebsite
		DIM	oRating
		DIM	oLogo
		DIM	oDisabled
				
		SET oID = RS.Fields("BrandID")
		SET	oName = RS.Fields("Name")
		SET oDesc = RS.Fields("Description")
		SET oWebsite = RS.Fields("Website")
		SET oRating = RS.Fields("Rating")
		SET oLogo = RS.Fields("Logo")
		SET	oDisabled = RS.Fields("Disabled")


%>

<form method="POST" action="brandeditsubmit.asp" enctype="multipart/form-data" name="EditForm" onsubmit="return validateForm()">
<input type="hidden" name="BrandID" value="<%=oID%>">
<div align="center">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
          <td align="right">Name</td>
          <td>
  <input type="text" name="Name" size="50" maxlength="50" value="<%=recString(oName)%>" class="required"></td>
        </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td>
  <textarea rows="5" name="Desc" cols="40"><%=recString(oDesc)%></textarea></td>
  </tr>
  <tr>
  <td align="right" valign="top">Website</td>
  <td>
  <input type="text" name="Website" size="40" value="<%=recString(oWebsite)%>"></td>
  </tr>
  <tr>
  <td align="right" valign="top">Rating</td>
  <td>
  <input type="text" name="Rating" size="5" value="<%=recNumber(oRating)%>"></td>
  </tr>
  <tr>
  <td align="right" valign="top">Logo</td>
  <td><input type="file" name="PictureFile" size="30">
<%

	DIM	sPictureName
	sPictureName = recString(oLogo)
	IF "" <> sPictureName THEN
		Response.Write "<br>"
		Response.Write "<img src=""picture.asp?table=brands&name=" & sPictureName & """><br>" & vbCRLF
		Response.Write "<input type=""checkbox"" name=""PictureDelete"" value=""ON""> Delete Logo" & vbCRLF
	END IF

%>
  </td>
  </tr>
  <tr>
  <td></td>
      <td>
  <input type="checkbox" name="Disabled" value="ON" <%=isTrue(oDisabled)%>>Disabled</td>
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
		END IF

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