<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT




DIM	g_sPage
g_sPage = Request("cat")

DIM	sMon
DIM	m
sMon = Request("m")
IF 0 < LEN(sMon) THEN
	m = CINT(sMon)
ELSE
	m = 3
END IF

DIM	sPreface
sPreface = Request("p")

DIM g_sNavigateTabLabel
g_sNavigateTabLabel = "(tab|special)"


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
Response.Write "<title>" & Server.HTMLEncode("Calendar - " & g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Calendar Print</title>
<%
END IF
%>
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" type="text/css" href="newsletter.css">
<link rel="stylesheet" type="text/css" href="<%=g_sTheme%>style.css">
<link rel="stylesheet" type="text/css" href="<%=remindCSS()%>">
<style type="text/css">
<!--

a:link, a:visited
{
	text-decoration: none;
}

.RemindMonthHeader
{
	border-bottom: 2px solid #000066;
	page-break-before: always;
}

tr:first-child th.RemindMonthHeader
{
	page-break-before: auto !important;
}




.RemindWeekRule
{
	height: 1px;
	border-bottom: 1px solid #99CCFF;
}

-->
</style>
</head>

<body>

<%

'===========================================================


DIM	oCalendar


DIM	nYear, nYearEnd
DIM	nMon, nMonEnd

DIM	dNow
dNow = DATEADD( "h", g_nServerTimeZoneOffset, NOW )

nYear = YEAR(dNow)
nMon = MONTH(dNow)
nMonEnd = nMon + m
IF 12 < nMonEnd THEN
	nYearEnd = nYear + FIX(nMonEnd/12)
	nMonEnd = nMonEnd MOD 12
ELSE
	nYearEnd = nYear
END IF

'Response.Write "m = " & nMon & ", y = " & nYear & "<br>"
'Response.Write "m = " & nMonEnd & ", y = " & nYearEnd & "<br>" & vbCRLF
		
DIM	nDateBegin
DIM	nDateEnd

nDateBegin = jdateFromGregorian( 1, nMon, nYear )
nDateEnd = jdateFromGregorian( 1, nMonEnd, nYearEnd ) - 1



'nDateBegin = jdateFromVBDate( NOW )
'nDateEnd = nDateBegin + 28 * 3 - 1
'nDateEnd = -14	'pending

gHtmlOption_encodeEmailAddresses = TRUE

DIM	sSavePictureFunc
sSavePictureFunc = g_htmlFormat_pictureFunc
g_htmlFormat_pictureFunc = "remindPicture"


SET oCalendar = loadRemindFiles2( nDateBegin, nDateEnd, _
							g_sPage, "holiday,usa,none", FALSE, dNow, "list" )
IF NOT Nothing IS oCalendar THEN

	IF "group" = sPreface  AND  FALSE THEN
%>
<!--webbot bot="Include" U-Include="_private/OurGroupRides.htm" TAG="BODY" startspan -->

		<ul style="margin-left: 1em; padding-left: 0;">
			<li>Unless otherwise indicated, all rides will leave from the ride 
			meeting place at the scheduled time.</li>
			<li>As a courtesy to all members, it is important to be ready to 
			leave at the departure time.</li>
			<li>All bikes should be fueled and bladders emptied ahead of time.</li>
			<li>We plan a variety of rides. Time and money might make it hard to 
			make all events. We hope this variety will give you an opportunity 
			to pick and choose, and at least enjoy some of them.</li>
			<li>If the <a href="links.asp?page=weather.htm">weather</a> is bad--come in your car.</li>
			<li>If you have a CB, you will find us on Channel One.</li>
		</ul>
		
<!--webbot bot="Include" i-checksum="1350" endspan -->
<%
	END IF

	'Load the XSL
	DIM	oXSL
	DIM	oXML
	SET oXSL = remindLoadXmlFile( Server.MapPath("scripts/remind.xslt") )
	'SET oXSL = Server.CreateObject("msxml2.DOMDocument")
	'oXSL.async = false
	'oXSL.load(Server.MapPath("scripts/remind.xslt"))
	
	SET oXML = oCalendar.xmldom
	Response.Write oXML.transformNode(oXSL)

END IF


%> 

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
