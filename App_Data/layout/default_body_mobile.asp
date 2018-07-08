<%
''<head>
''<title>1-Column (D)</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="dynamic">
''</head>

%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tableoutercolset">
<tr>
<td id="coldynamic">
<%
'========================
' ColDynamic
'========================


''1
''<head>
''<title>Calendar Next Meeting</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Notice our your next group meeting">
''</head>


SUB makeCalendarNextMeeting()

	DIM	nDateBegin
	DIM	nDateEnd

	nDateBegin = jdateFromVBDate( NOW )
	nDateEnd = nDateBegin + 7 * 12

Response.Write "<div class=""calendarcontent"">" & vbCRLF
Response.Write "<div class=""BlockHead"">Next Monthly Gathering</div>" & vbCRLF
Response.Write "<div class=""BlockBody"">"

outputDynamic "calendarnextmeeting", _
"loadcalendar.asp?f=calendarnextmeeting.htm&i=h&v=24&r=d&b=" & nDateBegin & "&e=" & nDateEnd & "&c=" & LCASE(g_sSiteChapterID) & "-gathering" & "&h=ignore&t=false&x=scripts/remind_gathering.xslt", _
		"calendarnextmeeting.htm", "0.5", "h", 24, "d"

Response.Write "</div>" & vbCRLF
Response.Write "</div>" & vbCRLF


END SUB
makeCalendarNextMeeting

''<head>
''<title>Calendar Events</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Upcoming events for your group">
''</head>


SUB makeCalendarUpcomingEvents()

	DIM	nDateBegin, nDateEnd

	nDateBegin = jdateFromVBDate( NOW )
	nDateEnd = nDateBegin + 7 * 3 - 1


Response.Write "<div class=""calendarcontent"">" & vbCRLF
Response.Write "<div class=""BlockHead""><a href=""calendar.asp?cat=" & LCASE(g_sSiteChapter) & """>Chapter Events / Rides</a></div>" & vbCRLF
Response.Write "<div class=""BlockBody"">"

outputDynamic "calendarupcoming", _
"loadcalendar.asp?f=calendarupcoming.htm&i=h&v=24&r=d&b=" & nDateBegin & "&e=" & nDateEnd & "&c=" & LCASE(g_sSiteChapter) & ",core,key" & "&h=holiday,usa,none&t=true&x=scripts/remind.xslt", _
		"calendarupcoming.htm", "0.5", "h", 24, "d"

Response.Write "</div>" & vbCRLF
Response.Write "</div>" & vbCRLF


END SUB
makeCalendarUpcomingEvents

''<head>
''<title>Forum Local</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Summary of recent forum postings for your group">
''</head>


RSSForum "forum.xml", g_sDomain, "localforum", g_sSiteTitle & " Forum", 20

''<head>
''<title>Weather Forecast</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Weather forecast for the zipcode provided in the site configuration">
''</head>

IF "" <> g_sSiteZip THEN
%>
<div id="weather" class="dynamic">
	<div class="BlockHead">Weather</div>
	<div class="BlockBody">
<%
outputDynamic "weather", "loadrssweather.asp?z=" & g_sSiteZip, "weather" & g_sSiteZip & ".htm", "1", "h", 2, "d"
%>
	</div>
</div>
<%
END IF




%>
</td>
</tr>
</table>
