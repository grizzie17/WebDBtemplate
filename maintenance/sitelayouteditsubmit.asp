<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True

DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%

DIM	sConfigFile
DIM	sOutFile
DIM	sLayout
DIM	sColumn1
DIM	sColumn2
DIM	sColumn3

sConfigFile = Request("configfile")
sOutFile = Request("outfile")
sLayout = Request("layout")
sColumn1 = Request("column-1")
sColumn2 = Request("column-2")
sColumn3 = Request("column-3")

IF "" = sConfigFile THEN sConfigFile = "config.txt"
IF "" = sOutFile THEN sOutFile = "default_body.asp"

DIM	sFolderLayout
	sFolderLayout = findAppDataFolder( "layout\homepage" )

DIM	sFolderConfig
	sFolderConfig = g_oFSO.BuildPath( sFolderLayout, "config" )

	createFolderHierarchy sFolderConfig


	DIM	oFile
	SET oFile = g_oFSO.CreateTextFile( sFolderConfig & "\" & sConfigFile )
	IF NOT oFile IS Nothing THEN
		oFile.WriteLine "layout:" & sLayout
		oFile.WriteLine "column-1:" & sColumn1
		oFile.WriteLine "column-2:" & sColumn2
		oFile.WriteLine "column-3:" & sColumn3
		oFile.Close
		SET oFile = Nothing
	END IF

	DIM	aCol(3)
	aCOl(0) = sColumn1
	aCol(1) = sColumn2
	aCol(2) = sColumn3
	DIM	aData(3)
	aData(0) = "''1"
	aData(1) = "''2"
	aData(2) = "''3"

	DIM	sFolderComponents
	sFolderComponents = findAppDataFolder( "layout\homepage\components" )

	DIM	aFiles
	DIM	sFile
	DIM	sData
	DIM	sCol
	DIM	i
	' load up the data for each column
	FOR i = LBOUND(aCol) TO UBOUND(aCol)
		sCol = aCol(i)
		IF 0 < LEN(sCol) THEN
			aFiles = SPLIT(sCol&",", ",")
			FOR EACH sFile IN aFiles
				IF 0 < LEN(sFile) THEN
					sFile = sFolderComponents & "\" & sFile
					IF g_oFSO.FileExists( sFile ) THEN
						SET oFile = g_oFSO.OpenTextFile(sFile)
						IF NOT oFile IS Nothing THEN
							sData = oFile.ReadAll
							sData = MID(sData,INSTR(sData,"'"))
							aData(i) = aData(i) & vbCRLF & sData
							oFile.Close
							SET oFile = Nothing
						END IF
					END IF
				END IF
			NEXT
		END IF
	NEXT

	DIM	aLayouts
	aLayouts = SPLIT(sLayout, ",")
	IF 0 < UBOUND(aLayouts) THEN
		sFile = CSTR(UBOUND(aLayouts)+1) & "Column"
		FOR EACH sCol IN aLayouts
			sFile = sFile & LEFT(sCol,1)
		NEXT
		sFile = sFile & ".txt"
	ELSE
		sFile = "1ColumnD.txt"
	END IF


	DIM	sFolderLayouts
	sFolderLayouts = findAppDataFolder( "layout\homepage\layouts" )
	DIM	sTemplate
	sTemplate = ""
	sFile = sFolderLayouts & "\" & sFile
	IF g_oFSO.FileExists( sFile ) THEN
		SET oFile = g_oFSO.OpenTextFile( sFile )
		IF NOT oFile IS Nothing THEN
			sTemplate = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
		END IF
	END IF

	FOR i = LBOUND(aLayouts) TO UBOUND(aLayouts)
		sCol = aLayouts(i)
		sTemplate = REPLACE(sTemplate,"%%"&sCol&"%%", aData(i))
	NEXT

	sFolderLayouts = findAppDataFolder("layout")
	sFile = sFolderLayouts & "\" & sOutFile
	IF g_oFSO.FileExists(sFile) THEN
		g_oFSO.DeleteFile(sFile)
	END IF
	SET oFile = g_oFSO.CreateTextFile( sFile )
	IF NOT oFile IS Nothing THEN
		oFile.Write sTemplate
		oFile.Close
		SET oFile = Nothing
	END IF

	cache_clearFolders


SET g_oFSO = Nothing

Response.Redirect "site.asp"

%>