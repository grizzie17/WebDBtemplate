<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<title>Upgrade DB</title>
<style type="text/css">
<!--

pre
{
	font-size: xx-small;
}



-->
</style>
</head>

<body>

<%


FUNCTION isSpace( s )
	isSpace = FALSE
	IF " " = LEFT(s, 1)  OR  vbTAB = LEFT(s, 1) THEN
		isSpace = TRUE
	END IF
END FUNCTION


FUNCTION TrimCRLF( sScript )

	DIM	i
	DIM	s
	s = TRIM(sScript)
	
	DO WHILE vbCRLF = LEFT(s,2)  OR  "#" = LEFT(s,1)  OR  isSpace( s )
		IF "#" = LEFT(s,1) THEN
			i = INSTR(s, vbCRLF)
			IF 0 < i THEN
				s = MID(s,i)
			ELSE
				s = ""
			END IF
		ELSEIF isSpace( s ) THEN
			s = TRIM( s )
		ELSE
			s = TRIM(MID(s,3))
		END IF
	LOOP
	
	DO WHILE vbCRLF = RIGHT(s,2)
		s = TRIM(LEFT(s,LEN(s)-2))
	LOOP
	TrimCRLF = s

END FUNCTION


DIM	sSelect
sSelect = "" _
	&	"SELECT " _
	&	"	* " _
	&	"FROM " _
	&	"	[Version] " _
	&	";"

DIM	oRS
ON Error Resume Next
SET oRS = dbQueryRead( g_DC, sSelect )
ON Error Goto 0

IF NOT oRS IS Nothing THEN
	DIM	nMajor
	DIM	nMinor
	DIM	nFix
	nMajor = 0
	nMinor = 0
	nFix = 0
	IF 0 < oRS.RecordCount THEN
		
		nMajor = recNumber( oRS.Fields("Major") )
		nMinor = recNumber( oRS.Fields("Minor") )
		IF 0 < nMajor THEN
			nFix = recNumber( oRS.Fields("Fixes") )
		ELSE
			nFix = 0
		END IF
	ELSE
	
		sSelect = "" _
			&	"INSERT INTO Version ( Major, Minor ) VALUES ( 0, 0 ); " _
			&	" "
		g_DC.Execute sSelect,,adCmdText+adExecuteNoRecords
	END IF
		
		
	'now get list of files from dbscripts folder
	
	DIM	i
	DIM	oFolder
	DIM	oFile
	DIM	sName
	DIM	aFileList()
	REDIM aFileList(10)
	DIM	nFileCount
	nFileCount = 0
	
	SET oFolder = g_oFSO.GetFolder( Server.MapPath( "./dbscripts" ) )
	
	FOR EACH oFile IN oFolder.Files
		sName = LCase( oFile.Name )
		i = INSTR(sName, ".txt")
		IF 0 < i THEN
			IF UBOUND(aFileList) <= nFileCount THEN
				REDIM PRESERVE aFileList(nFileCount+20)
			END IF
			aFileList(nFileCount) = LCASE(LEFT(sName,i-1)) & vbTAB & oFile.Path
			nFileCount = nFileCount + 1
		END IF
	NEXT 'oFile
	
	IF 0 < nFileCount THEN
		Response.Write "<h2>Processing " & nFileCount & " Upgrade Scripts</h2>" & vbCRLF
		Response.Flush
		sort aFileList, 0, nFileCount-1
		REDIM PRESERVE aFileList(nFileCount-1)
		
		DIM	sThisVers
		sThisVers = RIGHT("00" & nMajor,3) & "-" & RIGHT("00" & nMinor,3)
		IF 0 < nFix THEN
			sThisVers = sThisVers & "-" & RIGHT("00" & nFix,3)
		END IF
		'sThisVers = sThisVers
		
		DIM	sScript
		DIM	sReadScript
		DIM	aScripts
		DIM	aSplit
		DIM	sRec
		DIM	bProcess
		bProcess = FALSE
		FOR EACH sRec IN aFileList
			aSplit = SPLIT(sRec,vbTAB)
			IF NOT bProcess THEN
				IF aSplit(0) = sThisVers THEN
					bProcess = TRUE
				END IF
			END IF
			IF bProcess THEN
				SET oFile = g_oFSO.OpenTextFile( aSplit(1), 1 )
				IF NOT oFile IS Nothing THEN
					sReadScript = oFile.ReadAll
					oFile.Close
					SET oFile = Nothing
					aScripts = SPLIT(sReadScript,";")
					Response.Write "<p>Processing: " & aSplit(0) & "</p>" & vbCRLF
					Response.Flush
					FOR EACH sScript IN aScripts
						sScript = TrimCRLF( sScript )
						IF "" <> sScript THEN
							sScript = sScript & ";"
							Response.Write "<pre>" & sScript & "</pre><br>"
							Response.Flush
							g_DC.Execute sScript,,adCmdText+adExecuteNoRecords
						END IF
					NEXT 'sScript
				ELSE
					Response.Write "<p>Problem getting upgrade script " & aSplit(0) & "</p>" & vbCRLF
					EXIT FOR
				END IF
			END IF
		NEXT 'sRec
		IF NOT bProcess THEN
			Response.Write "<p>Database is already up-to-date.</p>" & vbCRLF
		END IF



	ELSE
		Response.Write "<p>No DB Upgrade scripts found</p>" & vbCRLF
	END IF
	oRS.Close
	SET oRS = Nothing
ELSE
		sSelect = "" _
			&	"CREATE TABLE Version (" _
			&	"	[RID] AUTOINCREMENT CONSTRAINT pk_RID PRIMARY KEY, " _
			&	"	[Major] LONG, " _
			&	"	[Minor] LONG " _
			&	"); " _
			&	"" _
			&	"INSERT Version ( Major, Minor ) VALUES ( 0, 0 ); "
		g_DC.Execute sSelect,,adCmdText+adExecuteNoRecords
	Response.Write "<p>Problem reading version info from database</p>" & vbCRLF
END IF



%>

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
