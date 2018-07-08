<%
'look below functions to find config declarations



FUNCTION getDomainRoot()
	DIM	sServer
	DIM	sPath
	DIM	aPath
	DIM	sPart
	DIM	i
	sPath = ""
	sServer = Request.ServerVariables("SERVER_NAME")
	IF sServer <> "localhost" THEN
		aPath = SPLIT(sServer,".")
		FOR i = LBOUND(aPath) TO UBOUND(aPath)-1	' -1 to dispose of tld
			sPart = aPath(i)
			IF "www" <> LCASE(sPart) THEN
				sPath = sPath & sPart
			END IF
		NEXT
		sPath = Replace( sPath, "-", "" )
		'getCookiePrefix = sPath
	ELSE
		sPath = Server.MapPath( "/" )
		aPath = SPLIT( sPath, ":" )
		sPath = aPath(1)
		aPath = SPLIT( sPath, "\" )
		FOR i = LBOUND(aPath) TO UBOUND(aPath)
			IF "wwwroot" = LCASE(aPath(i)) THEN
				aPath(i) = ""
				EXIT FOR
			ELSEIF i+1 < UBOUND(aPath) THEN
				aPath(i) = ""
			END IF
		NEXT 'i
		sPath = REPLACE(REPLACE(TRIM(JOIN( aPath, "" )), " ", "" ), "-", "" )
		'Response.Write "CookiePrefix = " & sPath & "<br>"
		'Response.Flush
	END IF
	getDomainRoot = sPath
	
END FUNCTION

'FUNCTION config_dateFromWeekNumber( nWeek, nWDay, nMonth, nYear )
'	DIM	dFirst
'	DIM	nFirstDay
'	DIM	x
'	dFirst = DATEVALUE(nMonth & "/1/" & nYear )
'	nFirstDay = WEEKDAY(dFirst)
'	x = 1 + ( nWDay - nFirstDay + 7 ) MOD 7
'	x = x + (7 * (nWeek - 1))
'	config_dateFromWeekNumber = DATEVALUE( nMonth & "/" & x & "/" & nYear )
'	'Response.Write "dateFromWeekNumber = " & config_dateFromWeekNumber & "<br>"
'END FUNCTION
'
'SUB config_adjustDaylightSavingsTime()
'	DIM	nMonth
'	DIM	nYear
'	nMonth = MONTH(NOW)
'	nYear = YEAR(NOW)
'	IF 3 < nMonth  AND  nMonth < 11 THEN
'		g_nServerTimeZoneOffset = g_nServerTimeZoneOffset + 1	' Daylight Savings Time
'	ELSEIF 3 = nMonth THEN
'		'second sunday
'		IF 0 < DATEDIFF("h",config_dateFromWeekNumber(2, 1, 3, nYear), NOW) THEN
'			g_nServerTimeZoneOffset = g_nServerTimeZoneOffset + 1	' Daylight Savings Time
'		END IF
'	ELSEIF 11 = nMonth THEN
'		'first sunday
'		IF 0 > DATEDIFF("h",config_dateFromWeekNumber(1, 1, 11, nYear), NOW) THEN
'			g_nServerTimeZoneOffset = g_nServerTimeZoneOffset + 1	' Daylight Savings Time
'		END IF
'	END IF
'	
'END SUB






DIM	g_sSiteName
g_sSiteName = "Undefined Site Name"

DIM g_sShortSiteName
g_sShortSiteName = "Udf Short Name"

DIM g_sSiteTitle
g_sSiteTitle = g_sSiteName

DIM g_sSiteMotto
g_sSiteMotto = ""

DIM g_sSiteChapter
g_sSiteChapter = "SS-CC"

DIM g_sSiteChapterID
g_sSiteChapterID = "CC"

DIM g_sSiteCity
g_sSiteCity = "Unknown, SS"

DIM	g_sSiteZip
g_sSiteZip = ""

DIM	g_sChapterNeighbors
g_sChapterNeighbors = ""



DIM	g_sRSSAnnounceInclude
g_sRSSAnnounceInclude = ""

DIM g_sRSSAnnounceExclude
g_sRSSAnnounceExclude = "plaque,personal"




DIM	g_nCopyrightStartYear
	g_nCopyrightStartYear = 0
DIM g_sCopyright
g_sCopyright = "&copy; " & DatePart("yyyy",NOW) & "&nbsp;  " & g_sShortSiteName & "&nbsp; All rights reserved"


DIM	g_sDomainRoot
g_sDomainRoot = getDomainRoot()

''' database

DIM	g_sDBProvider
	g_sDBProvider = "SQLOLEDB"

DIM	g_sDBServer
	g_sDBServer = "mssql301.ixwebhosting.com"

DIM	g_sDBSource
	g_sDBSource = LCASE( g_sDomainRoot )

DIM	g_sDBUser
	g_sDBUser = "user"

DIM	g_sDBPassword
	g_sDBPassword = "password"

DIM	g_sDBAdmin
	g_sDBAdmin = "admin"

DIM	g_sDBAdminPassword
	g_sDBAdminPassword = "password"

DIM	g_bDBDSN
	g_bDBDSN = false

DIM g_sDBConnectString
	g_sDBConnectString = ""


DIM g_sCookiePrefix
g_sCookiePrefix = LEFT( g_sDomainRoot, 8 )



DIM	g_sDomain
g_sDomain = "Undefined.tld"


DIM	g_sMailServer
g_sMailServer = "mail.Undefined.tld"

DIM g_sAnnonUser
DIM g_sAnnonPW

g_sAnnonUser = "notifications@" & LCASE(g_sDomain)
g_sAnnonPW = "bear1701"


DIM g_sSiteTabAnnouncements
g_sSiteTabAnnouncements = "Extra!"

DIM g_sSiteTabNewsletters
g_sSiteTabNewsletters = "Newsletters"

DIM g_sHomepageAnnouncements
g_sHomepageAnnouncements = g_sSiteTabAnnouncements

DIM g_sHomepageLayout
g_sHomepageLayout = "ace"
	'  ace - announce, calendar, extras
	'  eca
	'  cae
	'  cea
	'  aec
	'  eac


DIM	g_sDistrictDomain
g_sDistrictDomain = Request.ServerVariables("HOST_NAME")


DIM	g_sMetaKeywords
	g_sMetaKeywords = "gwrra, goldwing, gold wing road riders association, honda, motorcycle"




DIM g_bMailPickup
g_bMailPickup = FALSE



' UpcomingEvents
DIM g_sUpcomingEventsSchedule
'	mo,fr		indicates every monday and friday
g_sUpcomingEventsSchedule = "mo,fr"

DIM g_UE_nDurationCalendar
g_UE_nDurationCalendar = 6

DIM g_UE_nDurationForum
g_UE_nDurationForum = 8

DIM g_UE_nDurationAnnounce
g_UE_nDurationAnnounce = 6



DIM	g_nPictureDelay
g_nPictureDelay = 7		'seconds





dim g_bCalendarPrefaceOption
g_bCalendarPrefaceOption = TRUE


DIM g_sCalendarList
g_sCalendarList = ""

DIM g_sCalendarHiddenList
g_sCalendarHiddenList = ""



DIM	g_sLiveTileGathering
g_sLiveTileGathering = ""

DIM	g_sLiveTileEvent
g_sLiveTileEvent = ""




DIM g_nServerTimeZoneOffset
DIM	g_nTimeZoneOffset




g_nTimeZoneOffset = -6
IF LCASE(Request.ServerVariables("SERVER_NAME")) = "localhost" THEN
	g_nServerTimeZoneOffset = 0
ELSE
	g_nServerTimeZoneOffset = -6
	g_nServerTimeZoneOffset = 0
	'config_adjustDaylightSavingsTime
END IF





%>
<!--#include virtual="/config/configuser.asp"-->