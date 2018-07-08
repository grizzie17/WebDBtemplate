<%@	Language=VBScript
	EnableSessionState=True %>
<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2014 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------

OPTION EXPLICIT

DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )
%>
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="text/javascript" language="JavaScript">
<!--
function replaceWindowURL( win, url )
{
	win.location.href = url;
}
//-->
</script>
<script type="text/javascript" language="JavaScript1.1">
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
DIM	sID
sID = Request("id")
IF 0 < LEN(sID) THEN
	DIM	oXML
	DIM	sFilename
	DIM	sRemindFile
	
	sFilename = Request("filename")
	IF 0 = LEN(sFilename) THEN sFilename = "remind.xml"

	sRemindFile = findRemindFile( sFilename )
	IF "" <> sRemindFile THEN
		SET oXML = remindLoadXmlFile( sRemindFile )
		IF NOT oXML IS Nothing THEN
			DIM	oReminders
			DIM	oOldEvent
			SET oOldEvent = Nothing
			SET oReminders = oXML.selectSingleNode("/reminders")
			SET oOldEvent = oXML.selectSingleNode("//event[@id=""" & sID & """]")
			oReminders.removeChild( oOldEvent )
			oXML.save sRemindFile
			
			SET oOldEvent = Nothing
			SET oReminders = Nothing
			SET oXML = Nothing
		END IF
	END IF
	
	DIM	dNow
	dNow = NOW
	DIM	sT
	sT = CSTR(YEAR(dNOW)) & RIGHT("0"&MONTH(dNow),2) & RIGHT("0"&DAY(dNow),2)
	sT = sT & RIGHT("0"&HOUR(dNow),2) & RIGHT("0"&MINUTE(dNow),2) & RIGHT("0"&SECOND(dNow),2)

END IF

cache_clearFolders
%>
<script type="text/javascript" language="JavaScript">
<!--
	replaceWindowURL( self, "remind.asp?t=<%=sT%>&filename=<%=sFilename%>" );
//-->
</script>
</body>

</html>