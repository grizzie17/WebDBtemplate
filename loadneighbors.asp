<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT



%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/remind_cache.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts/include_theme.asp"-->
<%






FUNCTION genCalendar()

	DIM	s
	s = ""

	DIM	dNow
	dNow = DATEADD( "h", g_nServerTimeZoneOffset, NOW)


	DIM	nDateBegin
	DIM	nDateEnd


	DIM	sRmdHtm


	nDateBegin = jdateFromVBDate( dNow )
	nDateEnd = nDateBegin + 7 * 8

	sRmdHtm = remind_cache( "homeneighbors.htm", "h", 24, "d", _
						nDateBegin, nDateEnd, "neighbor-meeting", "ignore", FALSE, dNow, _
						Server.MapPath("scripts/remind_neighbors.xslt") )

	IF "" <> sRmdHtm THEN
		
		s=s& "<div class=""BlockHead"">Neighbors</div>" & vbCRLF
		s=s& "<div class=""BlockBody"">" & sRmdHtm & "</div>" & vbCRLF

	ELSE

		s=s&	"<div></div>" & vbCRLF

	END IF



	genCalendar = s

END FUNCTION


FUNCTION getCalendar()

	DIM	sData
	sData = ""
	DIM	sCacheFile
	sCacheFile = "neighbors.htm"
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, "h", 24, "d" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sData = genCalendar()
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", sCacheFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
	getCalendar = sData
	
END FUNCTION


DIM	sData
sData = getCalendar()
IF 0 < LEN(sData) THEN
	Response.ContentType = "text/text"
	Response.Write sData
ELSE
	Response.Status = "404 Not Found"
END IF

Response.Flush
Response.End




%>
<!--#include file="scripts\include_db_end.asp"-->