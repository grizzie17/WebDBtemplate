<!--#include file="sortutil.asp"-->
<%



FUNCTION getBasePart( sPath )

	DIM	sPart
	DIM	i
	
	sPart = sPath
	i = INSTRREV(sPart,"\")
	IF 0 < i THEN
		sPart = MID(sPart,i+1)
	END IF
	
	i = INSTRREV(sPart,"." )
	IF 0 < i THEN
		sPart = LEFT(sPart,i-1)
	END IF
	
	getBasePart = sPart

END FUNCTION



DIM	g_newsletter_sDomain
g_newsletter_sDomain = ""

FUNCTION getNewsletterDomain()
	IF 0 = LEN(g_newsletter_sDomain) THEN
		g_newsletter_sDomain = LCASE(Request.ServerVariables("HTTP_HOST"))
		IF "www." = LEFT(g_newsletter_sDomain,4) THEN
			g_newsletter_sDomain = MID(g_newsletter_sDomain,5)
		END IF
	END IF
	getNewsletterDomain = g_newsletter_sDomain
END FUNCTION




DIM	g_inheritNewsletter_makeLinkFunc
g_inheritNewsletter_makeLinkFunc = g_htmlFormat_makeLinkFunc

FUNCTION newsletter_buildLink( sProtocol, sURL, sTarget, sText )
	IF "mailto:" = sProtocol THEN
		IF gHtmlOption_encodeEmailAddresses THEN
			DIM	i
			DIM	sTE
			sTE = sURL
			i = INSTR( sURL, "@" )
			IF 0 < i THEN
				DIM	sDomain
				sDomain = getNewsletterDomain()
				DIM	sRoot
				sRoot = LCASE(MID(sURL,i+1))
				IF sRoot = sDomain THEN
					sTE = LEFT(sURL,i-1)
				END IF
			END IF

			sTE = REPLACE( sTE, "@", "*" )
			sTE = REPLACE( sTE, ".", ":" )
			
			FOR i = 3 TO LEN(sTE) STEP 5
				sTE = LEFT( sTE, i ) & " " & MID( sTE, i+1 )
			NEXT 'i
			
			IF "" = sText  OR  sURL = sText THEN
				newsletter_buildLink = "<span class=""pobox"">" & sTE & "</span>"
			ELSE
				newsletter_buildLink = "<span class=""pobox"">" & Server.HTMLEncode(sText) & "[" & sTE & "]</span>"
			END IF
		ELSE
			newsletter_buildLink = "<a href=""mailto:" & sURL & """>" & Server.HTMLEncode(sText) & "</a>"
		END IF
	ELSE
		newsletter_buildLink = EVAL( g_inheritNewsletter_makeLinkFunc & "( sProtocol, sURL, sTarget, sText )" )
		'newsletter_buildLink = htmlFormat_f_makeLink( sProtocol, sURL, sTarget, sText )
	END IF
END FUNCTION


g_htmlFormat_makeLinkFunc = "newsletter_buildLink"




DIM	g_sNewletterFolder
DIM	g_sRootFolder

DIM g_dNewslettersLastModified
g_dNewslettersLastModified = CDATE("1/1/1990")



FUNCTION includeBodyFilePath( aRootSplit(), sPath )

	DIM	sResultPath
	sResultPath = ""
	IF "/" = LEFT(sPath,1)  OR  INSTR(1,sPath,"//",vbTextCompare)  _
			OR  "http:" = LEFT(sPath,5)  OR  "https:" = LEFT(sPath,6) _
			OR  "mailto:" = LEFT(sPath,7)  OR  "news:" = LEFT(sPath,5) _
			OR  "javascript:" = LEFT(sPath,11)  OR  "#" = sPath  THEN
		sResultPath = sPath
	ELSE
		IF g_sNewletterFolder <> "" THEN
			sResultPath = virtualFromPhysicalPath(g_sNewletterFolder & "\images\" & sPath)
			sResultPath = "picture.asp?file=" & Server.URLEncode(sResultPath)
		END IF
	END IF
	includeBodyFilePath = sResultPath

END FUNCTION

g_includebody_filePathFunc = "includeBodyFilePath"	'( aRootSplit(), sPath )


FUNCTION pagePicture( sLabel )

	IF g_sNewletterFolder <> "" THEN
		pagePicture = virtualFromPhysicalPath(g_sNewletterFolder & "\images\" & sLabel)
		pagePicture = "picture.asp?file=" & Server.URLEncode(pagePicture)
	'	pagePicture = MID(pagePicture, LEN(g_sRootFolder)+1)
	'	pagePicture = REPLACE(pagePicture, "\", "/")
	END IF

END FUNCTION

g_htmlFormat_pictureFunc = "pagePicture"

gHtmlOption_pullquoteWidth = 175
gHtmlOption_sidebarWidth = 200


FUNCTION findNewslettersFolder( sBase )
	DIM	sFolder
	sFolder = findAppDataFolder( sBase )
	g_sNewletterFolder = sFolder
	g_sRootFolder = Server.MapPath( "/" )
	findNewslettersFolder = sFolder
END FUNCTION


SUB getAllNewsletters( aLetters(), sBase )

	DIM	sFolder
	sFolder = findNewslettersFolder( sBase )
	
	IF "" <> sFolder THEN
		DIM	oFolder
		DIM	oFile
		DIM	sFile
		DIM	sSuffix
		DIM	dTemp
		DIM	i
		DIM	j
		
		SET oFolder = g_oFSO.GetFolder( sFolder )
		
		i = 0
		REDIM aLetters(5)
		FOR EACH oFile IN oFolder.Files
			sFile = LCASE(oFile.Name)
			j = INSTR( sFile, "." )
			IF 0 < j THEN
				sSuffix = MID( sFile, j )
			ELSE
				sSuffix = ""
			END IF
			SELECT CASE sSuffix
			CASE ".txt", ".htm"
				IF UBOUND(aLetters) <= i THEN
					REDIM PRESERVE aLetters(i+5)
				END IF
				dTemp = oFile.DateLastModified
				IF g_dNewslettersLastModified < dTemp THEN g_dNewslettersLastModified = dTemp
				aLetters(i) = LCASE(oFile.Path)
				i = i + 1
			END SELECT
		NEXT 'oFile
		
		IF 0 < i THEN
			REDIM PRESERVE aLetters(i-1)
			sortDescend aLetters, 0, i-1
		END IF
		
		SET oFolder = Nothing
	END IF

END SUB




%>
