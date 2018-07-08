<%@	Language=VBScript
	EnableSessionState=True %>
<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2008 Bear Consulting Group
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
%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts/findfiles.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<%

DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )

'===========================================================

DIM	aFileList()
REDIM aFileList(5)
DIM	aSplit

getRemindList aFileList

DIM	sFilename

sFilename = Request("filename")
IF 0 = LEN(sFilename) THEN
	DIM	j
	aSplit = SPLIT(aFileList(0),vbTAB)
	j = INSTRREV(aSplit(1),"\")
	sFilename = MID(aSplit(1),j+1)
END IF


Response.Expires = 0

Response.ContentType = "text/html"


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<title>Calendar Maintenance</title>
<link rel="stylesheet" type="text/css" href="<%=remindCSS()%>">
<style type="text/css">
<!--

th
{
	background-color: #CCCCCC;
	font-size: x-small;
}

th.title
{
	background-color: #CCFFCC;
	font-size: medium;
}

.stylename
{
	font-size: x-small;
}



.categoryname
{
	font-size: small;
}



-->
</style>
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
<script type="text/javascript" language="JavaScript">
<!--

function selectItemOnValue( objOpt, sValue )
{
	var nMaxOptions;
	var	n;

	nMaxOptions = objOpt.length;
	for ( n = 0; n < nMaxOptions; ++n )
	{
		if ( sValue == objOpt.options[n].value )
		{
			objOpt.selectedIndex = n;
			return;
		}
	}
}

//-->
</script>
</head>

<body onload="window.focus()">

<form id="remindForm" name="remindForm">
<table border="0" width="100%" cellspacing="0" cellpadding="2">
  <tr>
    <th align="left" class="title" bgcolor="#CCFFCC">Calendar Maintenance</th>
    <td class="title" bgcolor="#CCFFCC"><select size="1" name="selectFile" onchange="options_select(this)">
<%
	DIM	sFile
	DIM	sTemp
	DIM	i
	FOR EACH sFile IN aFileList
		aSplit = SPLIT(sFile,vbTab)
		i = INSTRREV(aSplit(1),"\")
		sTemp = MID(aSplit(1),i+1)
		Response.Write "<option value=""" & sTemp & """>"
		sTemp = LEFT(sTemp,LEN(sTemp)-4)	' strip off suffix
		i = INSTR(sTemp,";")
		IF 0 < i THEN sTemp = LEFT(sTemp,i-1)
		sTemp = UCASE(LEFT(sTemp,1)) & MID(sTemp,2)
		sTemp = REPLACE( sTemp, "_", " " )
		Response.Write Server.HTMLEncode(sTemp) & "</option>" & vbCRLF
	NEXT 'sTemp
%>
      </select></td>
    <td align="right" class="title" bgcolor="#CCFFCC"><a href="javascript:window.close()">Done</a></td>
  </tr>
</table>
</form>

<script type="text/javascript" language="JavaScript">
<!--

function onFormLoad()
{
	var	oForm;
	var	oObj;
	
	oForm = document.remindForm;
	oObj = oForm.selectFile;
	selectItemOnValue( oObj, "<%=sFilename%>" )
}
onFormLoad();


function options_select(obj)
{
	var sUrl;
	var n = obj.selectedIndex;
	var sFile = obj.options[n].value;
	
	sUrl = "remind.asp?filename=" + sFile;
	replaceWindowURL( window, sUrl );
	return false;
}


//-->
</script>

<%

DIM	sRemindFile

sRemindFile = findRemindFile( sFilename )
IF 0 < LEN(sRemindFile) THEN

	DIM	oXSL
	DIM	oXML
	SET oXML = remindLoadXmlFile( sRemindFile )
	SET oXSL = remindLoadXmlFile( Server.MapPath("remind_list_edit.xslt") )
	
	DIM	oInc
	SET oInc = oXML.selectSingleNode("/reminders/include")
	IF oInc IS Nothing THEN
%>


<p><a href="remind_edit.asp?filename=<%=sFilename%>">Add New Event</a></p>

<%
	END IF
	'Transform the file
	'Response.Write oXML.transformNode(oXSL)
	DIM	s
	s = oXML.transformNode(oXSL)
	's = REPLACE( s, "_$$", "&" )
	Response.Write REPLACE(s, "$$Filename$$", sFilename, 1, -1, vbTextCompare)
	SET oXML = Nothing
	SET oXSL = Nothing
	IF NOT oInc IS Nothing THEN
		SET oXML = remindLoadWithInclude( sRemindFile )
		IF NOT oXML IS Nothing THEN
			SET oXSL = remindLoadXmlFile( Server.MapPath("remind_list.xslt") )
			IF NOT oXSL IS Nothing THEN
				s = oXML.transformNode( oXSL )
				Response.Write s
				SET oXSL = Nothing
			END IF
			SET oXML = Nothing
		END IF
	END IF
	SET oInc = Nothing
ELSE
	Response.Write "<p>Unable to locate " & sFilename & "</p>" & vbCRLF
END IF

%>

</body>

</html>