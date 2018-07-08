<%



DIM	g_sXDescription
DIM	g_sXBody
DIM	g_sXPicture
'g_sXDescription = ""
'g_sXBody = ""
'g_sXPicture = ""


DIM	g_inheritPagebody_makeLinkFunc

FUNCTION pagebody_buildLink( sProtocol, sURL, sTarget, sText )
	IF "mailto:" = sProtocol THEN
		IF gHtmlOption_encodeEmailAddresses THEN
			DIM	sTE
			DIM	i
			i = INSTR(sURL,"@")
			IF 0 < i THEN
				DIM	sTemp
				sTEMP = MID(sURL, i+1)
				IF LCASE(sTEMP) = LCASE(g_sDomain) THEN
					sTE = LEFT(sURL, i-1)
				ELSE
					sTE = sURL
				END IF
			ELSE
				sTE = sURL
			END IF
			sTE = REPLACE( sTE, "@", "*" )
			sTE = REPLACE( sTE, ".", ":" )
			
			FOR i = 3 TO LEN(sTE) STEP 5
				sTE = LEFT( sTE, i ) & " " & MID( sTE, i+1 )
			NEXT 'i
			
			IF "" = sText  OR  sURL = sText THEN
				pagebody_buildLink = "<span class=""pobox"">" & sTE & "</span>"
			ELSE
				pagebody_buildLink = "<span class=""pobox"">" & Server.HTMLEncode(sText) & "[" & sTE & "]</span>"
			END IF
		ELSE
			pagebody_buildLink = "<a href=""mailto:" & sURL & """>" & Server.HTMLEncode(sText) & "</a>"
		END IF
	ELSE
		pagebody_buildLink = EVAL( g_inheritPagebody_makeLinkFunc & "( sProtocol, sURL, sTarget, sText )" )
		'pagebody_buildLink = htmlFormat_f_makeLink( sProtocol, sURL, sTarget, sText )
	END IF
END FUNCTION







DIM	g_sPageImages
g_sPageImages = ""



FUNCTION pagebody_pagePicture( sPath )
	pagebody_pagePicture = ""

	DIM	oReg
	SET oReg = NEW RegExp
	oReg.Pattern = "^[\w\.-]+$"
	oReg.IgnoreCase = TRUE
	oReg.Global = TRUE
	IF oReg.Test(sPath) THEN
		pagebody_pagePicture = pagebodyPicture( sPath )
	ELSE
		pagebody_pagePicture = sPath
	END IF

END FUNCTION


FUNCTION includeBodyFilePath( aRootSplit(), sPath )

	DIM	sResultPath
	sResultPath = ""
	IF "/" = LEFT(sPath,1)  OR  INSTR(1,sPath,"//",vbTextCompare)  _
			OR  "http:" = LEFT(sPath,5)  OR  "https:" = LEFT(sPath,6) _
			OR  "mailto:" = LEFT(sPath,7)  OR  "news:" = LEFT(sPath,5) _
			OR  "javascript:" = LEFT(sPath,11)  OR  "#" = sPath  THEN
		sResultPath = sPath
	ELSE
		sResultPath = pagebody_pagePicture( sPath )
		IF "" = sResultPath THEN
			sResultPath = sPath
		END IF
	END IF
	includeBodyFilePath = sResultPath

END FUNCTION








DIM	g_oPictureDict
SET g_oPictureDict = Nothing

SUB loadLabelPictures()

	DIM	sSelect
	DIM	oRS
	
	SET g_oPictureDict = Server.CreateObject("Scripting.Dictionary")
	IF NOT Nothing IS g_oPictureDict THEN
	
		IF "" <> g_sXPicture THEN
			g_oPictureDict.Add "default", g_sXPicture
		END IF
		
		sSelect = "" _
			&	"SELECT " _
			&		"Label, " _
			&		"ImageID " _
			&	"FROM " _
			&		"pictures " _
			&	"WHERE " _
			&		"PageID = " & g_nPageID & " " _
			&	";"
		
		SET oRS = dbQueryRead( g_DC, sSelect )
		IF 0 < oRS.RecordCount THEN
	
			DIM oLabel
			DIM sLabel
			DIM oImageID
			DIM nImageID
			
			SET oImageID = oRS.Fields("ImageID")
			SET oLabel = oRS.Fields("Label")
			
			DO UNTIL oRS.EOF
				sLabel = LCASE(recString( oLabel ))
				nImageID = recNumber( oImageID )
				IF 0 < nImageID  AND  "" <> sLabel THEN
					'Response.Write "Load= " & sLabel & ", " & sFile & "<br>" & vbCRLF
					'Response.Flush
					g_oPictureDict.Add sLabel, nImageID
				END IF
				oRS.MoveNext
			LOOP
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
END SUB


FUNCTION pictureFileFromLibrary( sLbl )
	pictureFileFromLibrary = 0
	
	DIM	sSelect
	DIM	oRS
	DIM	sLib
	
	IF "~" = LEFT(sLbl,1) THEN
		sLib = MID(sLbl,2)
	ELSE
		sLib = sLbl
	END IF
	
	sSelect = "" _
		&	"SELECT " _
		&		"Label, " _
		&		"ImageID " _
		&	"FROM " _
		&		"(pictures " _
		&		"INNER JOIN pages ON pages.RID = pictures.PageID) " _
		&		"INNER JOIN categories ON categories.CategoryID = pages.Category " _ 
		&	"WHERE " _
		&		"categories.Name='@Library' " _
		&		"AND pages.PageName='~Library' " _
		&		"AND pictures.PageID = pages.RID " _
		&		"AND pictures.Label LIKE '" & sLib & "' " _
		&	";"
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF 0 < oRS.RecordCount THEN
	
		IF NOT Nothing IS g_oPictureDict THEN
		
			DIM oLabel
			DIM sLabel
			DIM oImageID
			DIM nImageID
			DIM	sTemp
			
			SET oImageID = oRS.Fields("ImageID")
			SET oLabel = oRS.Fields("Label")
			
			DO UNTIL oRS.EOF
				sLabel = "library-" & LCASE(recString( oLabel ))
				nImageID = recNumber( oImageID )
				IF 0 < nImageID  AND  "" <> sLabel THEN
					IF g_oPictureDict.Exists(sLabel) THEN
						sTemp = TRIM(g_oPictureDict.Item(sLabel))
						IF "" = sTemp THEN
							g_oPictureDict.Remove(sLabel)
							g_oPictureDict.Add sLabel, nImageID
							'Response.Write "Replaced= @ " & sLabel & ", " & sFile & "<br>" & vbCRLF
							'Response.Flush
							pictureFileFromLibrary = nImageID
						ELSE
							pictureFileFromLibrary = CLNG(sTemp)
						END IF
					ELSE
						'Response.Write "Add= @ " & sLabel & ", " & sFile & "<br>" & vbCRLF
						'Response.Flush
						g_oPictureDict.Add sLabel, nImageID
						pictureFileFromLibrary = nImageID
					END IF
				END IF
				oRS.MoveNext
			LOOP
		END IF
	END IF
	oRS.Close
	SET oRS = Nothing
END FUNCTION


FUNCTION pictureFileFromLabel( sLabel )
	pictureFileFromLabel = 0

	DIM	nImageID
	nImageID = 0
	
	IF Nothing IS g_oPictureDict THEN loadLabelPictures
	
	IF "~" = LEFT(sLabel,1) THEN
		DIM	s
		s = "library-" & LCASE(MID(sLabel,2))
		s = g_oPictureDict.Item( s )
		IF ISNUMERIC(s) THEN nImageID = CLNG(s)
		IF 0 = nImageID THEN
			pictureFileFromLabel = pictureFileFromLibrary( MID(sLabel,2) )
		ELSE
			pictureFileFromLabel = nImageID
		END IF
	ELSE
		pictureFileFromLabel = g_oPictureDict.Item( LCASE(sLabel) )
	END IF
	
END FUNCTION



FUNCTION pagebodyPicture( sLabel )

	DIM	nImageID
	nImageID = 0
	
	nImageID = pictureFileFromLabel( sLabel )
	IF 0 < nImageID THEN
		pagebodyPicture = "picture.asp?id=" & nImageID
	ELSE
		pagebodyPicture = ""
	END IF
END FUNCTION

FUNCTION pagebodyDownload( sLabel )

	DIM	nImageID
	nImageID = 0
	
	nImageID = pictureFileFromLabel( sLabel )
	IF 0 < nImageID THEN
		pagebodyDownload = "download.asp?id=" & nImageID & "&label=" & sLabel
	ELSE
		pagebodyDownload = ""
	END IF
END FUNCTION




DIM g_pagebody_htmlFormat_pictureFunc
DIM g_pagebody_htmlFormat_downloadFunc
DIM g_pagebody_htmlFormat_makeLinkFunc
DIM	g_pagebody_includebody_filePathFunc

SUB pagebody_saveFuncs()

	g_pagebody_htmlFormat_pictureFunc = g_htmlFormat_pictureFunc
	g_pagebody_htmlFormat_downloadFunc = g_htmlFormat_downloadFunc
	g_pagebody_htmlFormat_makeLinkFunc = g_htmlFormat_makeLinkFunc
	g_pagebody_includebody_filePathFunc = g_includebody_filePathFunc

	g_inheritPagebody_makeLinkFunc = g_htmlFormat_makeLinkFunc
	g_htmlFormat_makeLinkFunc = "pagebody_buildLink"
	g_htmlFormat_pictureFunc = "pagebodyPicture"
	g_htmlFormat_downloadFunc = "pagebodyDownload"
	g_includebody_filePathFunc = "includeBodyFilePath"	'( aRootSplit(), sPath )

END SUB

SUB pagebody_restoreFuncs()
	g_htmlFormat_pictureFunc = g_pagebody_htmlFormat_pictureFunc
	g_htmlFormat_downloadFunc = g_pagebody_htmlFormat_downloadFunc
	g_htmlFormat_makeLinkFunc = g_pagebody_htmlFormat_makeLinkFunc
	g_includebody_filePathFunc = g_pagebody_includebody_filePathFunc
END SUB




FUNCTION pagebody_formatBody( sFormat, sBody )

	DIM	sSaveEmailOption
	sSaveEmailOption = gHtmlOption_encodeEmailAddresses
	gHtmlOption_encodeEmailAddresses = TRUE
	
	SELECT CASE sFormat
	CASE "STFT"
		pagebody_formatBody = htmlFormatCRLF( sBody )
	CASE "4MTX"
		pagebody_formatBody = htmlFormatForum( sBody )
	CASE "HTML"
		pagebody_formatBody = includeBodyStreamFromString( sBody )
	CASE ELSE
		pagebody_formatBody = Server.HTMLEncode( sBody )
	END SELECT
	
	gHtmlOption_encodeEmailAddresses = sSaveEmailOption

END FUNCTION

FUNCTION pagebody_formatString( sFormat, sString )

	DIM	sSaveEmailOption
	sSaveEmailOption = gHtmlOption_encodeEmailAddresses
	gHtmlOption_encodeEmailAddresses = TRUE
	
	SELECT CASE sFormat
	CASE "STFT"
		pagebody_formatString = htmlFormatString( sString )
	CASE "4MTX"
		pagebody_formatString = htmlFormatForumString( sString )
	CASE "HTML"
		pagebody_formatString = sString
	CASE ELSE
		pagebody_formatString = Server.HTMLEncode( sString )
	END SELECT
	
	gHtmlOption_encodeEmailAddresses = sSaveEmailOption

END FUNCTION




SUB outputFormattedPageInfo( sFormat, sDescription, sBody, sPicture )



	DIM sText
	
	IF "" <> sBody THEN
		sText = pagebody_formatBody( sFormat, sBody )
		Response.Write sText
	ELSEIF "" <> sDescription THEN
		sText = pagebody_formatString( sFormat, sDescription )
		Response.Write sText
	END IF




END SUB





%>