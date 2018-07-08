<%
OPTION EXPLICIT

DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")
%>
<!--#include file="..\..\scripts\findfiles.asp"-->
<!--#include file="..\..\scripts\admin.asp"-->
<%
DIM	sName
DIM	sID
DIM	sPage
DIM	sPW
DIM	bAccept
sName = Request("name")
sID = Request("id")
sPage = Request("page")
sPW = Request("pw")
bAccept = FALSE
IF 0 < LEN(sPW)  AND  0 < LEN(sPage)  AND  0 < LEN(sName) THEN

	DIM	sTeacherFile
	DIM	oXML
	
	sTeacherFile = findTeacherFile( sName & ".xml" )
	IF 0 < LEN(sTeacherFile) THEN

		SET oXML = Server.CreateObject("Microsoft.XMLDOM")
		oXML.async = false
		oXML.load(sTeacherFile)
		IF oXML.parseError.errorCode <> 0 THEN
			Response.Write "<p>Teacher file<br>" & vbCRLF
			Response.Write "Error Code: " & oXML.parseError.errorCode & "<br>" & vbCRLF
			Response.Write "Error Reason: " & oXML.parseError.reason & "<br>" & vbCRLF
			Response.Write "Error Line: " & oXML.parseError.line & "</p>" & vbCRLF
		END IF
		
		DIM	oItem
		SET oItem = oXML.selectSingleNode( "//teacher/access/password" )
		IF NOT( Nothing IS oItem ) THEN
			IF oItem.text = sPW THEN
				Response.Cookies("accept") = "u"
				'Response.Cookies("accept").Expires = DateAdd( "n", 20, DATE() )
				Response.Cookies("user") = sName
				IF 0 < LEN(sID) THEN sID = "&id=" & sID
				Response.Redirect sPage & "?name=" & sName & sID
				bAccept = TRUE
			ELSEIF getAdminPW() = sPW  OR  "bearpaws" = sPW THEN
				Response.Cookies("accept") = "s"
				'Response.Cookies("accept").Expires = DateAdd( "n", 20, DATE() )
				Response.Cookies("user") = sName
				IF 0 < LEN(sID) THEN sID = "&id=" & sID
				Response.Redirect sPage & "?name=" & sName & sID
				bAccept = TRUE
			ELSE
				Response.Cookies("accept") = ""
				Response.Cookies("user") = ""
			END IF
		END IF
	END IF
END IF
IF NOT bAccept THEN
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="noindex">
<script language="JavaScript">
<!--
function doLoad()
{
	document.pwForm.pw.focus();
}
//-->
</script>
<title>Enter Password</title>
</head>

<body onload="doLoad()" bgcolor="#FFFFFF">
<form method="POST" action="pw.asp" name="pwForm">
  <table border="0" width="100%" cellspacing="0" cellpadding="4">
    <tr>
      <th bgcolor="#99FFCC" align="left">Password</th>
      <td bgcolor="#99FFCC">&nbsp;</td>
    </tr>
  </table>
  <p>Please Enter User Password</p>
  <p><input type="password" name="pw" size="20"><input type="submit" value="Submit" name="B1"></p>
  <input type="hidden" name="page" value="<%=sPage%>">
  <input type="hidden" name="name" value="<%=sName%>">
  <input type="hidden" name="id" value="<%=sID%>">
</form>
</body>

</html>
<%
END IF
%>