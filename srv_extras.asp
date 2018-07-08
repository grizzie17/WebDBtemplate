<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts/findfiles.asp"-->
<!--#include file="scripts/cookiesenabled.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts/htmlformat.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<%

DIM g_sHumorFilename
DIM	g_sHumorFolder
DIM	g_nHumorIndexCount
DIM g_dHumorIndexTime
DIM g_sHumorCookie
g_sHumorFilename = ""
g_sHumorFolder = ""
g_sHumorCookie = ""
g_nHumorIndexCount = 0
g_dHumorIndexTime = CDATE("1/1/2000")

DIM	g_sHTTP
g_sHTTP = "http://" & Request.ServerVariables("HTTP_HOST")

DIM	g_sPrefix
DIM	g_sHeader
g_sPrefix = LCASE(Request("q"))
g_sHeader = Server.HTMLEncode(Request("h"))
IF "" = g_sPrefix THEN
	g_sPrefix = "humor"
END IF
IF "" = g_sHeader THEN
	SELECT CASE g_sPrefix
	CASE "humor"
		g_sHeader = "Smile"
	CASE "safety"
		g_sHeader = "Safety Tips"
    CASE "quotes"
        g_sHeader = "Quotes"
	CASE ELSE
		g_sHeader = "---"
	END SELECT
END IF



FUNCTION findShuffleIndex( sBaseName )

	findShuffleIndex = ""

	DIM	sFolder

	sFolder = findAppDataFolder( sBaseName )
	IF "" <> sFolder THEN
	
		g_sHumorFolder = sFolder

		DIM	sIndexName
		sIndexName = cache_filepath( "shuffle", sBaseName&".txt" )
		'sIndexName = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
		
		IF g_oFSO.FileExists( sIndexName ) THEN
			findShuffleIndex = sIndexName
		END IF
	
	END IF

END FUNCTION


FUNCTION findHumorIndex()
	findHumorIndex = findShuffleIndex( g_sPrefix )
END FUNCTION


FUNCTION getNextCookie( sBaseName, nLow, nHigh )

	DIM	nCookie
	DIM sCookie
	DIM	sCookieName
	sCookieName = g_sCookiePrefix & "_" & sBaseName & "_index"
	sCookie = Request.Cookies( sCookieName )
	IF "" <> sCookie  AND  ISNUMERIC(sCookie) THEN
		nCookie = CINT(sCookie)
		nCookie = nCookie + 1
		IF nHigh < nCookie THEN
			nCookie = nLow
		ELSEIF nCookie < nLow THEN
			nCookie = nLow
		END IF
	ELSE
		DIM j
		RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
		nCookie = ROUND( RND * ( nHigh - nLow )) + nLow
		IF nHigh < nCookie THEN
			nCookie = nHigh
		ELSEIF nCookie < nLow THEN
			nCookie = nLow
		END IF
	END IF
	Response.Cookies( sCookieName ) = nCookie
	Response.Cookies( sCookieName ).Expires = DateAdd( "d", 56, NOW )
	
	getNextCookie = nCookie
	
	g_sHumorCookie = nCookie

END FUNCTION


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
			shuffle aLetters, 0, i-1
			shuffle aLetters, 0, i-1
		END IF
		
		SET oFolder = Nothing
	END IF

END SUB



SUB buildTextListIndex( sFolder, sIndexName )

	DIM	aJokes()
	REDIM aJokes(50)
	DIM	sJokesStream
	
	readTextListFolder sFolder, aJokes
	
	sJokesStream = JOIN( aJokes, vbCRLF )

	IF g_oFSO.FileExists( sIndexName ) THEN
		ON Error Resume Next
		g_oFSO.DeleteFile sIndexName, TRUE
		ON Error GOTO 0
	END IF
	
	ON Error Resume Next
	DIM	oFile
	SET oFile = g_oFSO.CreateTextFile( sIndexName, TRUE )
	IF NOT Nothing IS oFile THEN
		oFile.Write sJokesStream
		oFile.Close
		SET oFile = Nothing
	END IF
	ON Error GOTO 0

END SUB



FUNCTION needToRebuildIndex( sBaseName )

	needToRebuildIndex = FALSE

	DIM	dIndexTime
	DIM	nIndexCount
	DIM	nExpire
	
	dIndexTime = CDATE(g_dHumorIndexTime)
	nIndexCount = CINT(g_nHumorIndexCount)
	nExpire = nIndexCount / 2.5
	IF 7*8 < nExpire THEN nExpire = 7*8
	IF nExpire < ABS(DATEDIFF( "d", dIndexTime, NOW )) THEN
		DIM	sFolder
		DIM	sIndex
		sFolder = g_sHumorFolder
		sIndex = cache_filepath( "shuffle", sBaseName&".txt" )
		'sIndex = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
		
		'Response.Write "<span style=""color:gray"">" & LCASE(LEFT(sBaseName,1)) & "</span>"
		'IF Response.Buffer THEN Response.Flush
		buildTextListIndex sFolder, sIndex
		needToRebuildIndex = TRUE
	END IF

END FUNCTION




FUNCTION getTextListFilename( sBaseName )

	getTextListFileName = ""

	DIM	sFile
	sFile = findShuffleIndex( sBaseName )
	IF "" <> sFile THEN
		DIM	oIndexFile
		SET oIndexFile = g_oFSO.GetFile( sFile )
		IF NOT Nothing IS oIndexFile THEN
			g_dHumorIndexTime = CDATE(oIndexFile.DateLastModified)
			
			DIM oStream
			SET oStream = oIndexFile.OpenAsTextStream( 1 )	' read only
			IF NOT Nothing IS oStream THEN
				DIM	xData
				DIM	aData
				xData = oStream.ReadAll
				oStream.Close
				SET oStream = Nothing
				aData = SPLIT( xData, vbCRLF )
				
				g_nHumorIndexCount = UBOUND(aData) - LBOUND(aData) + 1
				
				DIM	i
				i = getNextCookie( sBaseName, LBOUND(aData), UBOUND(aData) )
				IF -1 < i THEN
					DIM	sPath
					DIM	sFolder
					sFolder = g_sHumorFolder
					sFile = aData(i)
					sPath = g_oFSO.BuildPath( sFolder, sFile )
					IF g_oFSO.FileExists( sPath ) THEN
						getTextListFilename = sPath
					END IF
				END IF
			END IF
			SET oIndexFile = Nothing
		END IF
	END IF

END FUNCTION


FUNCTION getHumorFilename()
	getHumorFilename = getTextListFilename( g_sPrefix )
END FUNCTION

g_sHumorFilename = getHumorFilename()


FUNCTION jokePicture( sLabel )

	DIM	sFolder
	sFolder = g_sHumorFolder
	jokePicture = virtualFromPhysicalPath(g_oFSO.BuildPath(sFolder, "images\" & sLabel))
	jokePicture = g_sHTTP & "/picture_old.asp?file=" & Server.URLEncode(jokePicture)

END FUNCTION


FUNCTION processJoke()

	processJoke = ""

	DIM	oFile
	DIM	oF
	
	IF "" <> g_sHumorFilename THEN
	
		SET oF = g_oFSO.GetFile( g_sHumorFilename )
		
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
		processJoke = htmlFormatCRLF( sText )
		
		g_htmlFormat_pictureFunc = sSavePictureFunc
	
	END IF
	
END FUNCTION


DIM	sJoke
sJoke = processJoke()
IF "" <> sJoke THEN
	Response.Write "<div class=""BlockHead"">" & g_sHeader & "</div>" & vbCRLF
	Response.Write "<div class=""BlockBody"">" & sJoke & "</div>"
END IF
needToRebuildIndex( g_sPrefix )


SET g_oFSO = Nothing
%>
