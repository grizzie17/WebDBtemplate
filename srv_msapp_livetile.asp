<%@	Language=VBScript%>
<%
OPTION EXPLICIT


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/remind_cache.asp"-->
<%
DIM	g_sHost
	g_sHost = Request.ServerVariables("HTTP_HOST")


FUNCTION remind_handleSubsTextDummy( sText, nDate, nTime )
	remind_handleSubsTextDummy = htmlFormatEncodeText( sText )
END FUNCTION

FUNCTION dummy_buildLink( sProtocol, sURL, sTarget, sText )
	DIM	sLink
	sLink = sProtocol & sURL
	IF sLink = sText THEN
		IF "address:" = sProtocol THEN
			dummy_buildLink = REPLACE(sURL, "+", " ")
		ELSEIF "mailto:" = sProtocol THEN
			dummy_buildLink = sURL
		ELSE
			dummy_buildLink = sLink
		END IF
	ELSE
		dummy_buildLink = sText
	END IF
END FUNCTION

FUNCTION dummyPictureGen( sAlign, sBorder, sWidth, sColor, sCaption, sFile )
	dummyPictureGen = " "
END FUNCTION


FUNCTION dummyPictureFile( sLabel )
	dummyPictureFile = " "
END FUNCTION



FUNCTION getStyleParam( sName, sParams )
	getStyleParam = ""
	DIM	oReg
	SET oReg = NEW RegExp
	oReg.Pattern = "\s*" & sName & ":\s*([^;]+);"
	oReg.Global = TRUE
	oReg.IgnoreCase = TRUE
	DIM	oMatchList
	DIM	oMatch
	SET oMatchList = oReg.Execute( sParams )
	IF NOT oMatchList IS Nothing THEN
		FOR EACH oMatch IN oMatchList
			getStyleParam = TRIM(oMatch.Submatches(0))
			EXIT FOR
		NEXT
		SET oMatchList = Nothing
	END IF
	SET oReg = Nothing
END FUNCTION




g_htmlFormat_makeLinkFunc = "dummy_buildLink"

'g_remind_handleSubsText = "remind_handleSubsTextDummy"



FUNCTION getEvent( sCategories, sCache )
	DIM	dNow
	dNow = NOW

	DIM	nDateBegin
	DIM	nDateEnd
	nDateBegin = jdateFromVBDate( dNow )
	nDateEnd = nDateBegin + 7 * 12

	getEvent = remind_cache( sCache, "h", 12, "d", _
						nDateBegin, nDateEnd, sCategories, "ignore", FALSE, dNow, _
						Server.MapPath("scripts/remind_livetile_gathering.xslt") )
	'getEvent = Server.HTMLEncode( getEvent )
	getEvent = REPLACE( getEvent, "&amp;", "&", 1, -1, vbTextCompare  )
	getEvent = REPLACE( getEvent, "&nbsp;", " ", 1, -1, vbTextCompare  )
END FUNCTION

FUNCTION findImage( sPath, sName )
	findImage = ""

	DIM sFile
	sFile = g_oFSO.BuildPath( sPath, sName & ".png" )
	IF g_oFSO.FileExists( sFile ) THEN
		findImage = "http://" & g_sHost & virtualFromPhysicalPath( sFile )
		EXIT FUNCTION
	END IF

	sFile = g_oFSO.BuildPath( sPath, sName & ".gif" )
	IF g_oFSO.FileExists( sFile ) THEN
		findImage = "http://" & g_sHost & virtualFromPhysicalPath( sFile )
		EXIT FUNCTION
	END IF

END FUNCTION




FUNCTION tilePicture()
	DIM	s
	s = ""

	DIM	oFolder
	DIM	sPath
	
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))
	sPath = oFolder.Path
	SET oFolder = Nothing

	DIM	s70
	DIM	s150
	DIM	s310x150
	DIM	s310

	s70 = findImage( sPath, "favicon128" )
	s150 = findImage( sPath, "favicon270" )
	s310x150 = findImage( sPath, "favicon588x270" )
	s310 = findImage( sPath, "favicon588" )

	IF 0 < (LEN(s150) + LEN(s310x150) + LEN(s310)) THEN

		s=s &	"<?xml version=""1.0"" ?>" & vbCRLF
		s=s &	"<tile>" & vbCRLF
		s=s &	"<visual lang=""en-US"" version=""2"">" & vbCRLF
		IF 0 < LEN(s150) THEN
			s=s &	"<binding template=""TileSquare150x150Image"" fallback=""TileSquareImage"">" & vbCRLF
			s=s &	"<image id=""1"" src=""" & s150 & """ alt=""" & Server.HTMLEncode(g_sSiteName) & """ />" & vbCRLF
			s=s &	"</binding>" & vbCRLF
		END IF
		IF 0 < LEN(s310x150) THEN
			s=s &	"<binding template=""TileWide310x150Image"" fallback=""TileWideImage"">" & vbCRLF
			s=s &	"<image id=""1"" src=""" & s310x150 & """ alt=""" & Server.HTMLEncode(g_sSiteName) & """ />" & vbCRLF
			s=s &	"</binding>" & vbCRLF
		END IF
		IF 0 < LEN(s310) THEN
			s=s &	"<binding template=""TileSquare310x310Image"">" & vbCRLF
			s=s &	"<image id=""1"" src=""" & s310 & """ alt=""" & Server.HTMLEncode(g_sSiteName) & """ />" & vbCRLF
			s=s &	"</binding>" & vbCRLF
		END IF
		s=s &	"</visual>" & vbCRLF
		s=s &	"</tile>" & vbCRLF

	END IF

	tilePicture = s
END FUNCTION

FUNCTION tileGathering()
	DIM	s
	s = ""

	DIM	sCategory
	IF 0 < LEN(g_sLiveTileGathering) THEN
		sCategory = g_sLiveTileGathering
	ELSE
		sCategory = LCASE(g_sSiteChapterID) & "-gathering"
	END IF

	DIM	sGathering
	sGathering = getEvent(sCategory, "livetilegathering.txt")
	IF 0 < LEN(sGathering) THEN
		s=s &	"<?xml version=""1.0"" ?>" & vbCRLF
		s=s &	"<tile>" & vbCRLF
		s=s &	"<visual lang=""en-US"" version=""2"">" & vbCRLF
		s=s &	"<binding template=""TileSquare150x150Text04"" fallback=""TileSquareText04"">" & vbCRLF
		s=s &	"<text id=""1"">" & sGathering & "</text>" & vbCRLF
		s=s	&	"</binding>" & vbCRLF
		s=s &	"<binding template=""TileWide310x150Text04"" fallback=""TileWideText04"">" & vbCRLF
		s=s &	"<text id=""1"">" & sGathering & "</text>" & vbCRLF
		s=s	&	"</binding>" & vbCRLF
		s=s &	"<binding template=""TileSquare310x310TextList02"">" & vbCRLF
		s=s &	"<text id=""1"">" & sGathering & "</text>" & vbCRLF
		s=s	&	"</binding>" & vbCRLF
		s=s &	"</visual>" & vbCRLF
		s=s &	"</tile>" & vbCRLF
	END IF

	tileGathering = s
END FUNCTION

FUNCTION tileEvent()
	DIM	s
	s = ""

	DIM	sCategory
	IF 0 < LEN(g_sLiveTileEvent) THEN
		sCategory = g_sLiveTileEvent
	ELSE
		sCategory = LCASE(g_sSiteChapter)
	END IF

	DIM	sGathering
	sGathering = getEvent(sCategory, "livetileevent.txt")
	IF 0 < LEN(sGathering) THEN
		s=s &	"<?xml version=""1.0"" ?>" & vbCRLF
		s=s &	"<tile>" & vbCRLF
		s=s &	"<visual lang=""en-US"" version=""2"">" & vbCRLF
		s=s &	"<binding template=""TileSquare150x150Text04"" fallback=""TileSquareText04"">" & vbCRLF
		s=s &	"<text id=""1"">" & sGathering & "</text>" & vbCRLF
		s=s	&	"</binding>" & vbCRLF
		s=s &	"<binding template=""TileWide310x150Text04"" fallback=""TileWideText04"">" & vbCRLF
		s=s &	"<text id=""1"">" & sGathering & "</text>" & vbCRLF
		s=s	&	"</binding>" & vbCRLF
		s=s &	"<binding template=""TileSquare310x310TextList02"">" & vbCRLF
		s=s &	"<text id=""1"">" & sGathering & "</text>" & vbCRLF
		s=s	&	"</binding>" & vbCRLF
		s=s &	"</visual>" & vbCRLF
		s=s &	"</tile>" & vbCRLF
	END IF

	tileEvent = s
END FUNCTION

FUNCTION getTile( sName, sFunc )
	DIM	sXML
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sName, "h", 12, "d" )
	IF NOT oFile IS Nothing THEN
		sXML = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sXML = EVAL( sFunc & "()" )
		IF 0 < LEN(sXML) THEN
			SET oFile = cache_makeFile( "site", sName )
			IF NOT oFile IS Nothing THEN
				oFile.Write sXML
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
	getTile = sXML
END FUNCTION

	DIM	sQ
	sQ = Request("q")

	DIM sXML

	SELECT CASE sQ
	CASE 1		' picture
		sXML = getTile( "livetile1.xml", "tilePicture" )
	CASE 2		' gathering
		sXML = getTile( "livetile2.xml", "tileGathering" )
	CASE 3		' chapter-event (non-gathering)
		sXML = getTile( "livetile3.xml", "tileEvent" )
	CASE ELSE
		sXML = getTile( "livetile1.xml", "tilePicture" )
	END SELECT

	IF 0 < LEN(sXML) THEN

		Response.ContentType = "text/xml"
		Response.Write sXML

	END IF


%>
<!--#include file="scripts\include_db_end.asp"-->
