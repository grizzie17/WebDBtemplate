<%
OPTION EXPLICIT
%>
<!--#include file="..\..\scripts\findfiles.asp"-->
<%

DIM g_oFSO
SET g_oFSO = CreateObject( "Scripting.FileSystemObject" )

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="noindex">
<script language="JavaScript">
<!--
function replaceWindowURL( win, url )
{
	win.location.href = url;
}
//-->
</script>
<script language="JavaScript1.1">
<!--
function replaceWindowURL( win, url )
{
	win.location.replace( url );
}
//-->
</script>
</head>

<body>
<%
DIM	sName
DIM	sTeacherName
DIM	oXML
DIM	sTeacherFile

sName = Request("name")
sTeacherName = sName
IF 0 < LEN(sName) THEN
	
	sTeacherFile = findTeacherFile( sName & ".xml" )
	IF 0 < LEN(sTeacherFile) THEN

		SET oXML = Server.CreateObject("Microsoft.XMLDOM")
		'SET oXML = Server.CreateObject("Msxml2.DOMDocument")
		oXML.async = false
		oXML.load(sTeacherFile)
		IF oXML.parseError.errorCode <> 0 THEN
			Response.Write "<p>Teacher file<br>" & vbCRLF
			Response.Write "Error Code: " & oXML.parseError.errorCode & "<br>" & vbCRLF
			Response.Write "Error Reason: " & oXML.parseError.reason & "<br>" & vbCRLF
			Response.Write "Error Line: " & oXML.parseError.line & "</p>" & vbCRLF
		END IF
		
		DIM	sPW
		
		sTeacherName = Request("teachername")
		sPW = Request("pw1")

		DIM	oNewPW
		DIM	oText
		DIM	oNode
		
		SET oText = oXML.createCDATASection( sPW )
		SET oNewPW = oXML.createNode( 1, "password", "" )
		oNewPW.appendChild( oText )
				
		DIM	oAccess
		DIM	oOldPW
		SET oAccess = oXML.documentElement.selectSingleNode("/teacher/access")
		SET oOldPW = oAccess.selectSingleNode("password")
		
		SET oNode = oAccess.replaceChild( oNewPW, oOldPW )
		
		oXML.save sTeacherFile

	END IF
END IF
%>
<script language="JavaScript">
<!--
	replaceWindowURL( self, "teacheredit.asp?name=<%=sName%>" );
//-->
</script>
</body>

</html>
