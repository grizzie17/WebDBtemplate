<%
''<head>
''<title>2-Column (DF)</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="dynamic,fixed1">
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
''<title>Forum Local</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Summary of recent forum postings for your group">
''</head>


RSSForum "forum.xml", g_sDomain, "localforum", g_sSiteTitle & " Forum", 20

''<head>
''<title>Announcements Local</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Announcements from your local announcements (Extra!) category">
''</head>

makeAnnouncements

''<head>
''<title>Homepage Articles</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Short articles to be included on your homepage">
''</head>


Response.Write "<div id=""homebodycontent"">" & vbCRLF
pagebody_saveFuncs
outputMultiplePages TRUE	' TRUE to generate an enclosing table
pagebody_restoreFuncs
Response.Write "</div>" & vbCRLF

''<head>
''<title>Quotes - Women</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Inspired and funny quotes from women">
''</head>


outputDynamic "quotes_women", "loadextras.asp?q=quotes&t=15&h=Quotes", "quotes_women.htm", "1.5", "n", 15, "d"




%>
</td>
<td id="colfixed1">
<%
'========================
' ColFixed1
'========================


''2
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
''<title>Calendar Events 2</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Upcoming events for your group">
''</head>


SUB makeCalendarUpcomingEvents2()

	DIM	nDateBegin, nDateEnd

	nDateBegin = jdateFromVBDate( NOW )
	nDateEnd = nDateBegin + 7 * 3 - 1

	DIM	sCategories
	sCategories = REPLACE(TRIM(LCASE(g_sCalendarCategories)), " ", "")
	IF "" = sCategories THEN
		sCategories = LCASE(g_sSiteChapter) & ",core,key"
	END IF

Response.Write "<div class=""calendarcontent"">" & vbCRLF
Response.Write "<div class=""BlockHead""><a href=""calendar.asp?cat=" & sCategories & """>Upcoming Events</a></div>" & vbCRLF
Response.Write "<div class=""BlockBody"">"

outputDynamic "calendarupcoming2", _
"loadcalendar.asp?f=calendarupcoming2.htm&i=h&v=24&r=d&b=" & nDateBegin & "&e=" & nDateEnd & "&c=" & sCategories & "&h=holiday,usa,none&t=true&x=scripts/remind.xslt", _
		"calendarupcoming2.htm", "0.5", "h", 24, "d"

Response.Write "</div>" & vbCRLF
Response.Write "</div>" & vbCRLF


END SUB
makeCalendarUpcomingEvents2

''<head>
''<title>Calendar Last Modified</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Display the date/time the calendar was last updated">
''</head>


IF ISDATE( dRemindLastModified ) THEN

%>
<p style="color:#999999; font-family: sans-serif; font-size: xx-small;">Calendar 
Updated: <%=DATEADD("h", g_nServerTimeZoneOffset, dRemindLastModified)%></p>
<%

END IF

''<head>
''<title>Calendar Neighbors</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Notices of your neighbors next group meetings">
''</head>


outputDynamic "neighbors", "loadneighbors.asp", "neighbors.htm", "2", "h", 24, "d"

''<head>
''<title>Calendar Birthdays</title>
''<meta name="navigate" content="tab">
''<meta name="description" content="Monthly birthday and anniversary announcements">
''</head>

'loadcalendar.asp?f=cachefile.htm&i=h&v=6&r=d&b=begindate&e=enddate&c=categories&h=holidaycats&t=showtoday&x=xsltfile

SUB makeCalendarBirthdays()

	DIM	nYear, nYearEnd
	DIM	nMon, nMonEnd
	DIM	nDateBegin, nDateEnd

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


Response.Write "<div class=""calendarcontent"">" & vbCRLF
Response.Write "<div class=""BlockHead""><a href=""calendar.asp?cat=" & sRmdCat & """>Birthdays / Anniversaries</a></div>" & vbCRLF
Response.Write "<div class=""BlockBody"">"

outputDynamic "calendarbirthday", _
"loadcalendar.asp?f=homebday.htm&i=d&v=7&r=m&b=" & nDateBegin & "&e=" & nDateEnd & "&c=" & sRmdCat & "&h=ignore&t=false&x=scripts/remind.xslt", _
		"homebday.htm", "0.5", "d", 7, "m"

Response.Write "</div>" & vbCRLF
Response.Write "</div>" & vbCRLF

END SUB
makeCalendarBirthdays



%>
</td>
</tr>
</table>
