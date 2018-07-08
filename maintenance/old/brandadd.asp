<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True



%>
<%




%><html>

<head>
<title>Add Brand</title>
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

<h1>Add Brand</h1>
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

<form method="POST" action="brandaddsubmit.asp" name="EditForm" onsubmit="return validateForm()">
<div align="center">
<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
          <td align="right">Name</td>
          <td>
  <input type="text" name="Name" size="40" maxlength="50" value="" class="required"></td>
        </tr>
  <tr>
  <td align="right" valign="top">Description</td>
  <td><textarea rows="7" name="Desc" cols="35"></textarea></td>
  </tr>
  <tr>
  <td align="right" valign="top">Website</td>
  <td><input type="text" name="Website" size="40"></td>
  </tr>
  <tr>
  <td align="right" valign="top">Rating</td>
  <td><input type="text" name="Rating" size="5"></td>
  </tr>
  </table>
  <p>
  <input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
  <input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
</div>

</form>

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
%>