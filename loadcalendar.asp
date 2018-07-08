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

FUNCTION ISBOOL( s )
	ISBOOL = FALSE
	DIM	t
	t = CSTR(s)
	IF 0 < LEN(t) THEN
		IF ISNUMERIC(t) THEN
			IF CLNG(t) <> 0 THEN ISBOOL = TRUE
		ELSE
			SELECT CASE LCASE(LEFT(t,1))
			CASE "t", "y"
				ISBOOL = TRUE
			END SELECT
		END IF
	END IF
END FUNCTION





FUNCTION genCalendar()

	DIM	s
	s = ""

	DIM	dNow
	dNow = DATEADD( "h", g_nServerTimeZoneOffset, NOW)


	DIM	nDateBegin
	DIM	nDateEnd


	DIM	sRmdHtm


	nDateBegin = jdateFromVBDate( dNow )
	nDateEnd = nDateBegin + 7 * 12

	sRmdHtm = remind_cache( "homegathering.htm", "h", 24, "d", _
						nDateBegin, nDateEnd, LCASE(g_sSiteChapterID) & "-gathering", "ignore", FALSE, dNow, _
						Server.MapPath("scripts/remind_gathering.xslt") )

	IF "" <> sRmdHtm THEN
		
		s=s& "<div class=""BlockHead"">Next Monthly Gathering</div>" & vbCRLF
		s=s& "<div class=""BlockBody"">" & sRmdHtm & "</div>" & vbCRLF

	END IF

	'===========================================================


	nDateBegin = jdateFromVBDate( dNow )
	nDateEnd = nDateBegin + 7 * 3 - 1

	sRmdHtm = remind_cache( "homeevents.htm", "h", 24, "d", _
						nDateBegin, nDateEnd, LCASE(g_sSiteChapter) & ",core,key", "holiday,usa,none", TRUE, dNow, _
						Server.MapPath("scripts/remind.xslt") )

	IF "" <> sRmdHtm THEN

		s=s& "<div class=""BlockHead""><a href=""calendar.asp?cat=" & LCASE(g_sSiteChapter) & """>Chapter Events / Rides</a></div>" & vbCRLF
		s=s& "<div class=""BlockBody"">" & sRmdHtm & "</div>" & vbCRLF
	END IF
	'===========================================================


	DIM	nYear, nYearEnd
	DIM	nMon, nMonEnd


	nYear = YEAR(NOW)
	nMon = MONTH(NOW)
	nMonEnd = nMon + 1
	IF 12 < nMonEnd THEN
		nYearEnd = nYear + FIX(nMonEnd/12)
		nMonEnd = nMonEnd MOD 12
	ELSE
		nYearEnd = nYear
	END IF

	nDateBegin = jdateFromGregorian( 1, nMon, nYear )
	nDateEnd = jdateFromGregorian( 1, nMonEnd, nYearEnd ) - 1


	DIM	sRmdCat
	sRmdCat = "birthday,anniversary"

	sRmdHtm = remind_cache( "homebday.htm", "d", 7, "m", _
						nDateBegin, nDateEnd, sRmdCat, "ignore", FALSE, dNow, _
						Server.MapPath("scripts/remind.xslt") )

	IF "" <> sRmdHtm THEN

		s=s& "<div class=""BlockHead""><a href=""calendar.asp?cat=" & sRmdCat & """>Birthdays / Anniversaries</a></div>" & vbCRLF
		s=s& "<div class=""BlockBody"">" & sRmdHtm & "</div>" & vbCRLF

	END IF

	IF ISDATE( dRemindLastModified ) THEN

		s=s&	"<p style=""color:#999999; font-family: sans-serif; font-size: xx-small;"">Calendar Updated: "
		s=s&	DATEADD("h", g_nServerTimeZoneOffset, dRemindLastModified) & "</p>" & vbCRLF

	END IF

	genCalendar = s

END FUNCTION


FUNCTION getCalendar(sCacheFile, qInterval, qIntValue, qBreakInterval, _
					qDateBegin, qDateEnd, qCategories, qHolidayCategories, qShowToday, dNow, _
					qXSLTFile)

	DIM	sData
	sData = ""
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, qInterval, qIntValue, qBreakInterval )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sData = remind_cache( sCacheFile, qInterval, qIntValue, qBreakInterval, _
					qDateBegin, qDateEnd, qCategories, qHolidayCategories, qShowToday, dNow, _
					Server.MapPath(qXSLTFile) )
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


DIM	qFile
DIM	qInterval
DIM	qIntValue
DIM	qBreakInterval
DIM	qDateBegin
DIM	qDateEnd
DIM	qCategories
DIM	qHolidayCategories
DIM	qShowToday
DIM	qToday
DIM	qXsltFile

	qFile = Request("f")
	qInterval = Request("i")
	qIntValue = CLNG(Request("v"))
	qBreakInterval = Request("r")
	qDateBegin = CLNG(Request("b"))
	qDateEnd = CLNG(Request("e"))
	qCategories = Request("c")
	qHolidayCategories = Request("h")
	qShowToday = CBOOL(Request("t"))
	qXsltFile = Request("x")

	DIM	dNow
	dNow = DATEADD( "h", g_nServerTimeZoneOffset, NOW)

	'loadcalendar.asp?f=cachefile.htm&i=h&v=6&r=d&b=begindate&e=enddate&c=categories&h=holidaycats&t=showtoday&x=xsltfile

DIM	sData
sData = getCalendar(qFile, qInterval, qIntValue, qBreakInterval, _
					qDateBegin, qDateEnd, qCategories, qHolidayCategories, qShowToday, dNow, _
					qXSLTFile)
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