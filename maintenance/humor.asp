<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>


<%






SUB readTextFolderObject( sRoot, oInFolder, aLetters(), nCountLetters )

	IF oInFolder IS Nothing THEN
		EXIT SUB
	END IF
	IF "_" = LEFT(oInFolder.Name,1) THEN
		EXIT SUB
	END IF
	
	DIM	oFolder
	DIM	oFile
	DIM	sFile
	DIM	i
	DIM	nLenRoot
	
	nLenRoot = LEN(sRoot)+1

	SET oFolder = oInFolder
			
	i = nCountLetters
	FOR EACH oFile IN oFolder.Files
		sFile = LCASE(oFile.Name)
		IF "_" <> LEFT(sFile,1) THEN
			IF 0 < INSTR(1, sFile, ".txt", vbTextCompare ) THEN
				IF UBOUND(aLetters) <= i THEN
					REDIM PRESERVE aLetters(i+50)
				END IF
				aLetters(i) = MID(oFile.Path,nLenRoot)
				i = i + 1
			END IF
		END IF
	NEXT 'oFile
	SET oFile = Nothing
	
	DIM oSubFolder
	FOR EACH oSubFolder IN oFolder.SubFolders
		IF "_" <> LEFT(oSubFolder.Name,1)  AND  "images" <> LCASE(oSubFolder.Name) THEN
			readTextFolderObject sRoot, oSubFolder, aLetters, i
		END IF
	NEXT 'oSubFolder
	
	nCountLetters = i
	


END SUB


SUB readTextListFolder( sFolder, aLetters() )

	
	IF "" <> sFolder THEN
		DIM	oFolder
		DIM	oFile
		DIM	sFile
		DIM sRoot
		DIM	i
		
		SET oFolder = g_oFSO.GetFolder( sFolder )
		
		sRoot = sFolder
		IF "\" <> RIGHT(sRoot,1) THEN
			sRoot = sRoot & "\"
		END IF
		
		i = 0
		REDIM aLetters(10)
		readTextFolderObject sRoot, oFolder, aLetters, i
		
		
		IF 0 < i THEN
			REDIM PRESERVE aLetters(i-1)
			sort aLetters, 0, i-1
		END IF
		
		SET oFolder = Nothing
	END IF

END SUB



SUB readJokesFolder( sFolder, aLetters() )

	
	IF "" <> sFolder THEN
	
		readTextListFolder sFolder, aLetters
		
	END IF

END SUB

	DIM aJokes()
	REDIM aJokes(10)

	DIM	sFolder
	sFolder = findAppDataFolder( "humor" )

	readJokesFolder sFolder, aJokes
	
	Response.Write "Total " & UBOUND(aJokes)+1
	
	Response.Write "<ul>" & vbCRLF
	
	DIM	sFile
	
	FOR EACH sFile IN aJokes
		Response.Write "<li>"
		Response.Write "<a target=""humorshow"" href=""humorshow.asp?file="
		Response.Write Server.URLEncode(g_oFSO.BuildPath(sFolder, sFile))
		Response.Write """>"
		Response.Write REPLACE(Server.HTMLEncode(sFile), "\", "&nbsp;&nbsp;\&nbsp;&nbsp;" )
		Response.Write "</a>"
		Response.Write "</li>" & vbCRLF
	NEXT 'sFile
	
	Response.Write "</ul>" & vbCRLF


%>


</body>

</html>
