<%@	Language=VBScript
	EnableSessionState=True %>
<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2014 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------

OPTION EXPLICIT

DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )


%>
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%



FUNCTION digit2( s )
	IF 1 < LEN( s ) THEN
		digit2 = s
	ELSE
		digit2 = "0" & s
	END IF
END FUNCTION

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script type="text/javascript" language="JavaScript">
<!--
function replaceWindowURL( win, url )
{
	win.location.href = url;
}
//-->
</script>
<script type="text/javascript" language="JavaScript1.1">
<!--
function replaceWindowURL( win, url )
{
	win.location.replace( url );
}
//-->
</script>
</head>

<body>
<%
DIM	sFilename
DIM	sID
DIM	sSubject
DIM	sContent
DIM	sLocation
DIM	sComments
DIM	sStyle
DIM	sCategories
DIM	sDate
DIM	sTime
DIM	sOffset
DIM	sDuration
DIM	sDurationView
DIM	sPending

DIM	sWeeklyRepeat
DIM	sWeeklyDays

DIM	sDateType
DIM	sDateSubType
DIM	sDD
DIM	sMM
DIM	sYY
DIM	sWD
DIM	sWN
DIM	sKey

sFilename = Request("filename")
IF 0 = LEN(sFilename) THEN sFilename = "remind.xml"

sID = Request("id")
sSubject = TRIM(Request("subject"))
sContent = TRIM(Request("content"))
sLocation = TRIM(Request("location"))
sComments = TRIM(Request("comments"))
sStyle = TRIM(Request("style"))
sCategories = TRIM(Request("categories"))
sTime = TRIM(Request("time"))
sOffset = TRIM(Request("offset"))
sDuration = TRIM(Request("duration"))
sDurationView = TRIM(Request("durationview"))
sPending = TRIM(Request("pending"))
sDateType = TRIM(Request("recurrence"))


IF "" <> sTime THEN
	IF LEN(sTime) < 5 THEN
		sTime = RIGHT("0000" & sTime, 5)
	END IF
END IF

SELECT CASE sDateType
CASE "single"
	sDD = Request("single_day")
	sMM = Request("single_month")
	sYY = Request("single_year")
	sDate = sYY & " " & digit2(sMM) & " " & digit2(sDD)
CASE "weekly"
	sWeeklyRepeat = Request("weeklyRepeat")
	sKey = ""
	sWD = Request("weeklySunday")
	IF "ON" = sWD THEN sKey = sKey & ",1"
	sWD = Request("weeklyMonday")
	IF "ON" = sWD THEN sKey = sKey & ",2"
	sWD = Request("weeklyTuesday")
	IF "ON" = sWD THEN sKey = sKey & ",3"
	sWD = Request("weeklyWednesday")
	IF "ON" = sWD THEN sKey = sKey & ",4"
	sWD = Request("weeklyThursday")
	IF "ON" = sWD THEN sKey = sKey & ",5"
	sWD = Request("weeklyFriday")
	IF "ON" = sWD THEN sKey = sKey & ",6"
	sWD = Request("weeklySaturday")
	IF "ON" = sWD THEN sKey = sKey & ",7"
	sKey = MID(sKey,2)
	IF "" = sWeeklyRepeat THEN sWeeklyRepeat = "0"
	IF "" = sKey THEN sKey = "1"
	sDate = sWeeklyRepeat & " " & sKey
CASE "monthly"
	sDateSubType = Request("radioMonthly")
	SELECT CASE sDateSubType
	CASE "dayn"
		sDD = Request("monthlyDayn_day")
		sMM = Request("monthlyDayn_index")
		sYY = Request("monthlyDayn_cycle")
		sDate = sYY & " " & sMM & " " & digit2(sDD)
	CASE "wday"
		sYY = Request("monthlyWDay_cycle")
		sMM = Request("monthlyWDay_index")
		sWN = Request("monthlyWDay_weekNumber")
		sWD = Request("monthlyWDay_weekday")
		sDate = sYY & " " & sMM & " " & sWN & " " & sWD
	END SELECT
CASE "yearly"
	sDateSubType = Request("radioYearly")
	SELECT CASE sDateSubType
	CASE "dayn"
		sMM = Request("yearlyDayn_month")
		sDD = Request("yearlyDayn_day")
		sDate = digit2(sMM) & " " & digit2(sDD)
	CASE "wday"
		sMM = Request("yearlyWDay_month")
		sWN = Request("yearlyWDay_weekNumber")
		sWD = Request("yearlyWDay_weekday")
		sDate = digit2(sMM) & " " & sWN & " " & sWD
	END SELECT
CASE "hebrew"
	sDateSubType = Request("radioHebrew")
	SELECT CASE sDateSubType
	CASE "dayn"
		sMM = Request("hebrewDayn_month")
		sDD = Request("hebrewDayn_day")
		sDate = digit2(sMM) & " " & digit2(sDD)
	CASE "wday"
		sMM = Request("hebrewWDay_month")
		sWN = Request("hebrewWDay_weekNumber")
		sWD = Request("hebrewWDay_weekday")
		sDate = digit2(sMM) & " " & sWN & " " & sWD
	END SELECT
CASE "keyword"
	sKey = Request("keyword_name")
	sDate = sKey
CASE "season"
	sKey = Request("season_name")
	sDate = sKey
END SELECT
%>
id = <%=sID%><br>
sSubject = <%=Server.HTMLEncode(sSubject)%><br>
sContent = <%=Server.HTMLEncode(sContent)%><br>
sLocation = <%=Server.HTMLEncode(sLocation)%><br>
sComments = <%=Server.HTMLEncode(sComments)%><br>
sStyle = <%=sStyle%><br>
sCategories = <%=sCategories%><br>
sDateType = <%=sDateType%><br>
sDateSubType = <%=sDateSubType%><br>
sDate = <%=sDate%><br>
sTime = <%=sTime%><br>
sOffset = <%=sOffset%><br>
sDuration = <%=sDuration%><br>
sDurationView = <%=sDurationView%>

<%

IF "" <> sDate  AND  "" <> sDateType THEN
	DIM	oXML
	DIM	sRemindFile

	sRemindFile = findRemindFile( sFilename )
	SET oXML = remindLoadXmlFile( sRemindFile )

		DIM	oEvent
		DIM	oDate
		DIM	oCategory
		DIM	oStyle
		DIM	oSubject
		DIM	oContent
		DIM oTime
		DIM	oText
		DIM	oAttr
		
		SELECT CASE LCASE(sDurationView)
		CASE "list"
			sDurationView = "list"
		CASE "mag"
			sDurationView = "mag"
		CASE "none"
			sDurationView = "none"
		CASE "all"
			sDurationView = ""
		CASE ELSE
			sDurationView = ""
		END SELECT
		
		SET oEvent = oXML.createNode( 1, "event", "" )
		IF 0 < LEN(sID) THEN
			oEvent.setAttribute "id", sID
		ELSE
			Randomize
			sKey = "e" & YEAR(NOW) & digit2(MONTH(NOW)) & digit2(DAY(NOW)) & digit2(HOUR(NOW)) & digit2(MINUTE(NOW)) & digit2(SECOND(NOW))
			oEvent.setAttribute "id", sKey
		END IF
		SET oText = oXML.createTextNode( sDate )
		SET oDate = oXML.createNode( 1, "date", "" )
		oDate.setAttribute "type", sDateType
		IF 0 < LEN(sDateSubType) THEN oDate.setAttribute sDateType, sDateSubType
		IF 0 < LEN(sTime) THEN oDate.setAttribute "time", sTime
		IF 0 < LEN(sOffset) THEN oDate.setAttribute "offset", sOffset
		IF 0 < LEN(sDuration) THEN oDate.setAttribute "duration", sDuration
		IF 0 < LEN(sDurationView) THEN oDate.setAttribute "duration-view", sDurationView
		IF 0 < LEN(sPending) THEN oDate.setAttribute "pending", sPending
		oDate.appendChild( oText )
		oEvent.appendChild( oDate )
		
		IF 0 < LEN( sCategories ) THEN
			PUBLIC aCatSplit
			DIM	sCat
			DIM	sCategory
			aCatSplit = SPLIT( sCategories, ",", -1, vbTextCompare )
			FOR EACH sCat IN aCatSplit
				sCategory = TRIM(sCat)
				IF 0 < LEN(sCategory) THEN
					SET oText = oXML.createTextNode( sCategory )
					SET oCategory = oXML.createNode( 1, "category", "" )
					oCategory.appendChild( oText )
					oEvent.appendChild( oCategory )
				END IF
			NEXT
		END IF
		
		IF 0 < LEN( sStyle ) THEN
			SET oText = oXML.createTextNode( sStyle )
			SET oStyle = oXML.createNode( 1, "style", "" )
			oStyle.appendChild( oText )
			oEvent.appendChild( oStyle )
		END IF
		'SET oText = oXML.createTextNode( sSubject )
		SET oText = oXML.createCDATASection( sSubject )
		SET oSubject = oXML.createNode( 1, "subject", "" )
		oSubject.appendChild( oText )
		oEvent.appendChild( oSubject )
		
		IF 0 < LEN( sLocation ) THEN
			SET oText = oXML.createCDATASection( sLocation )
			SET oContent = oXML.createNode( 1, "location", "" )
			oContent.appendChild( oText )
			oEvent.appendChild( oContent )
		END IF
		
		IF 0 < LEN( sContent ) THEN
			SET oText = oXML.createCDATASection( sContent )
			SET oContent = oXML.createNode( 1, "content", "" )
			oContent.appendChild( oText )
			oEvent.appendChild( oContent )
		END IF
		
		IF 0 < LEN( sComments ) THEN
			SET oText = oXML.createCDATASection( sComments )
			SET oContent = oXML.createNode( 1, "comments", "" )
			oContent.appendChild( oText )
			oEvent.appendChild( oContent )
		END IF
		

	DIM	oReminders
	DIM	oNewEvent
	DIM	oOldEvent
	SET oOldEvent = Nothing
	SET oReminders = oXML.selectSingleNode("/reminders")
	IF oReminders IS Nothing THEN
		Response.Write "something is VERY WRONG - unable to find root node<br>"
		Response.Flush
	END IF
	IF 0 < LEN(sID) THEN
		SET oOldEvent = oXML.selectSingleNode("//event[@id=""" & sID & """]")
	END IF
	IF Nothing IS oOldEvent THEN
		SET oNewEvent = oReminders.appendChild( oEvent )
	ELSE
		SET oNewEvent = oReminders.replaceChild( oEvent, oOldEvent )
	END IF
	SET oReminders = Nothing
		SET oEvent = Nothing
		SET oDate = Nothing
		SET oSubject = Nothing
		SET oContent = Nothing
		SET oText = Nothing
		
		oXML.save sRemindFile & "_t"
	SET oXML = Nothing
		
	g_oFSO.DeleteFile sRemindFile
	g_oFSO.CopyFile sRemindFile & "_t", sRemindFile, TRUE
	g_oFSO.DeleteFile sRemindFile & "_t"
		

	DIM	dNow
	dNow = NOW
	DIM	sT
	sT = CSTR(YEAR(dNOW)) & RIGHT("0"&MONTH(dNow),2) & RIGHT("0"&DAY(dNow),2)
	sT = sT & RIGHT("0"&HOUR(dNow),2) & RIGHT("0"&MINUTE(dNow),2) & RIGHT("0"&SECOND(dNow),2)

	cache_clearFolders

END IF
%>
<script type="text/javascript" language="JavaScript">
<!--
	replaceWindowURL( self, "remind.asp?t=<%=sT%>&filename=<%=sFilename%>" );
//-->
</script>
</body>

</html>
<%
	SET g_oFSO = Nothing
%>