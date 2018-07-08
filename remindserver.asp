<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<%

DIM	q_q
q_q = LCASE(Request("q"))


DIM	aList()
REDIM aList(10)

getRemindList aList

DIM	sLI
DIM	aItems
DIM	i
DIM	sFile
DIM	sFilename
DIM	sKey

sFilename = ""

FOR EACH sLI IN aList
	aItems = SPLIT(sLI,vbTAB)
	sFile = aItems(1)
	i = INSTRREV(sFile,"\")
	sKey = LCASE(MID(sFile,i+1))
	i = INSTRREV(sKey,".")
	sKey = LEFT(sKey,i-1)
	i = INSTR(sKey,";")
	IF 0 < i THEN
		sKey = LEFT(sKey,i-1)
	END IF
	IF q_q = sKey THEN
		sFilename = sFile
	END IF
NEXT 'sLI

IF "" <> sFilename THEN

	DIM	oFile
	SET oFile = g_oFSO.OpenTextFile( sFilename, 1 )
	IF NOT oFile IS Nothing THEN
	
		DIM	sData
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	
		Response.ContentType = "text/xml"
		Response.Write sData
		Response.Flush
		Response.End
	
	ELSE
	
		Response.Status = "404 Not Found"
	
	END IF

ELSE

	Response.Status = "403 Forbidden"
	Response.Write "403 Forbidden" & vbCRLF

END IF


SET g_oFSO = Nothing
%>