<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Server.ScriptTimeout = 60*30
Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin_admin.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<title>DB Upgrade</title>
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

	DIM	i, j
	DIM	s
	s = TRIM(sScript)
	
	DO WHILE vbCRLF = LEFT(s,2)  OR  "-- " = LEFT(s,3)  OR  isSpace( s )
		IF "-- " = LEFT(s,3) THEN
			i = INSTR(s, vbCRLF)
			IF 0 < i THEN
				s = MID(s,i+2)
			ELSE
				s = ""
			END IF
		ELSEIF vbCRLF = LEFT(s,2) THEN
			s = MID(s,3)
		END IF
		s = TRIM(s)
	LOOP
	
	DO WHILE vbCRLF = RIGHT(s,2)
		s = TRIM(LEFT(s,LEN(s)-2))
	LOOP
	TrimCRLF = s

END FUNCTION


DIM	sSelect
sSelect = "" _
	&	"IF NOT EXISTS " _
	&	"(SELECT * FROM sysobjects WHERE id = object_id(N'[version]') " _
	&	" AND OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
	&	" CREATE TABLE [version] ( " _
	&	"  [Major] INT NOT NULL, " _
	&	"  [Minor] INT NOT NULL, " _
	&	"  [Fixes] INT NULL DEFAULT NULL, " _
	&	" ) " _
	&	";"

DIM	oRS
SET oRS = dbExecute( g_DC, sSelect )
IF NOT oRS IS Nothing THEN

	SET oRS = Nothing



	sSelect = "" _
		&	"SELECT " _
		&	"	* " _
		&	"FROM " _
		&	"	[version] " _
		&	";"

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
				&	"INSERT INTO [version] ( Major, Minor ) VALUES ( 0, 0 ); " _
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
	
		SET oFolder = g_oFSO.GetFolder( Server.MapPath( "dbscripts/" & dbVendor() ) )
	
		FOR EACH oFile IN oFolder.Files
			sName = LCase( oFile.Name )
			i = INSTR(sName, ".sql")
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
			DIM	nErr
			DIM	sErr
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
						'sScript = CSTR(ASC(&hEF)) & CSTR(ASC(&hBB)) & CSTR(ASC(&hBF))
						'IF sScript = LEFT(sReadScript,3) THEN
							'dispose of UTF8 byte order mark
							sReadScript = MID(sReadScript,4)
						'END IF
						aScripts = SPLIT(sReadScript,";")
						Response.Write "<p>Processing: " & aSplit(0) & "</p>" & vbCRLF
						Response.Flush
						nErr = 0
						FOR EACH sScript IN aScripts
							'Response.Write "<pre>" & sScript & "</pre><br>"
							'Response.Flush
							sScript = TrimCRLF( sScript )
							IF "" <> sScript THEN
								sScript = sScript & ";"
								Response.Write "<pre>" & sScript & "</pre><br>"
								Response.Flush
								ON ERROR RESUME NEXT
								g_DC.Execute sScript,,adCmdText+adExecuteNoRecords
								nErr = Err.number
								sErr = Err.Description
								ON ERROR GOTO 0
								IF 0 <> nErr THEN
									IF &h80040E37 <> nErr THEN
										Response.Write "<p> -- Error = " & CSTR(HEX(nErr)) & " - " & sErr & "</p>" & vbCRLF
										EXIT FOR
									ELSE
										Response.Write "......<br>"
										nErr = 0
									END IF
									'i = INSTR( sScript, "DROP" )
									'IF 0 < i THEN
									'	nErr = 0
									'ELSE
									'	EXIT FOR
									'END IF
								END IF
							END IF
						NEXT 'sScript
						IF 0 <> nErr THEN
							EXIT FOR
						END IF
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

			ON ERROR RESUME Next
			sSelect = "" _
				&	"DROP TABLE [version];"
			g_DC.Execute sSelect ',,adCmdText+adExecuteNoRecords

			sSelect = "" _
				&	"CREATE TABLE [version] (" _
				&	"	[Major] INT NOT NULL, " _
				&	"	[Minor] INT NOT NULL, " _
				&	"	[Fixes] INT NULL DEFAULT NULL" _
				&	"); "

			g_DC.Execute sSelect ',,adCmdText+adExecuteNoRecords

			sSelect = "" _
				&	"INSERT INTO [version] ( [Major], [Minor] ) VALUES ( 0, 0 ); "
			g_DC.Execute sSelect ',,adCmdText+adExecuteNoRecords
			ON ERROR GOTO 0
		Response.Write "<p>Problem reading version info from database</p>" & vbCRLF
	END IF

ELSE
	Response.Write "<p>Problem establishing database</p>" & vbCRLF
END IF


%>

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
