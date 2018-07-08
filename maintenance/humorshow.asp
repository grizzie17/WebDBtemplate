<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\htmlformat.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Humor Show</title>
<link rel="stylesheet" type="text/css" href="../themes/default/style.css">
</head>

<body>

<%

DIM	g_sJokeFolder
g_sJokeFolder = findAppDataFolder( "humor" )



FUNCTION jokePicture( sLabel )

	DIM	sFolder
	sFolder = g_sJokeFolder
	jokePicture = virtualFromPhysicalPath(g_oFSO.BuildPath(sFolder, "images\" & sLabel))
	jokePicture = "picture.asp?file=" & Server.URLEncode(jokePicture)

END FUNCTION


DIM	sFile

sFile = Request("file")

	DIM	oFile
	DIM	oF
	
	SET oF = g_oFSO.GetFile( sFile )
	
	SET oFile = oF.OpenAsTextStream( 1 )	'open for read
	
	DIM	sText
	sText = oFile.ReadAll
	oFile.Close
	SET oFile = Nothing
	SET oF = Nothing

	DIM	sSavePictureFunc
	sSavePictureFunc = g_htmlFormat_pictureFunc
	g_htmlFormat_pictureFunc = "jokePicture"
	
	gHtmlOption_encodeEmailAddresses = TRUE
	
	DIM sFormat
	sFormat = htmlFormatCRLF( sText )
	
	g_htmlFormat_pictureFunc = sSavePictureFunc






%>
<table width="180" cellpadding="4">
<tr>
<td bgcolor="#FFEECC" class="SmileBody">
<%
Response.Write sFormat
%>
</td>
</tr>
</table>

</body>

</html>
