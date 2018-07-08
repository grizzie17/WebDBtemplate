<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<title>Clean Remind Files</title>
</head>

<body>


<%

DIM	aList()
REDIM aList(1)

getRemindList aList

DIM	oFile
DIM	sFile
DIM	sRemindFile
DIM	sData
DIM	sNewData
DIM	nStart
DIM	nEnd
DIM	nLen
DIM	n
DIM	aFileSplit
DIM	oMatch
DIM	oMatches
DIM	oReg
SET oReg = NEW RegExp

FOR EACH sFile IN aList
	'sRemindFile = findRemindFile( sFile )
	aFileSplit = SPLIT( sFile, vbTAB )
	sRemindFile = aFileSplit(1)
	Response.Write "<p>file: "  & Server.HTMLEncode(sRemindFile) & "</p>"
	Response.Flush
	SET oFile = g_oFSO.OpenTextFile( sRemindFile, 1 )
	IF NOT oFile IS Nothing THEN
		nStart = 1
		nEnd = 1
		sNewData = ""
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
		nLen = LEN(sData)
		IF 0 < nLen THEN
			n = 0
			oReg.Pattern = "<script>"
			oReg.Global = false
			DO
				SET oMatches = oReg.Execute( sData )
				n = 0
				FOR EACH oMatch IN oMatches
					n = n + 1
					nStart = oMatch.FirstIndex
					nEnd = INSTR( nStart, sData, "</script>" )
					IF 0 < nEnd THEN
						sData = LEFT( sData, nStart ) & MID( sData, nEnd+9 )
					END IF
				NEXT 'oMatch
			LOOP UNTIL n = 0
			sData = REPLACE( sData, "<bo"&"dy>", "<content>" )
			sData = REPLACE( sData, "</"&"bo"&"dy>", "</content>" )
			IF nLEN <> LEN(sData ) THEN
				SET oFile = g_oFSO.CreateTextFile( sRemindFile & "_t", true )
				IF NOT oFile IS Nothing THEN
					oFile.Write sData
					oFile.Close
					SET oFile = Nothing
					g_oFSO.CopyFile sRemindFile & "_t", sRemindFile, TRUE
					g_oFSO.DeleteFile sRemindFile & "_t"
				END IF
			END IF
		END IF
	END IF
NEXT 'sFile

SET oReg = Nothing
SET g_oFSO = Nothing


%>

</body>

</html>
