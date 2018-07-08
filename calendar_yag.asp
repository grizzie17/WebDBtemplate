<%@ Language=VBScript %>
<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2005 Bear Consulting Group
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


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")




DIM	q_year
DIM	q_startWeek
DIM q_daynameabbrev
DIM	q_startMonth
DIM	q_layout
DIM q_Progression
DIM	q_defWeekend

q_year = Request("year")
q_startWeek = Request("firstWeekday")
q_daynameabbrev = Request("daynameabbrev")
q_startMonth = Request("firstMonth")
q_layout = Request("layout")
q_Progression = Request("Progression")
q_defWeekend = Request("defWeekend")

%>
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts\remind.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN">
<!--DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%=q_year%></title>
<style type="text/css">


.yearpage
{
	text-align: center;
}


.yearblock
{
	font-family: sans-serif;
	text-align: center;
}

.yearblock td, .yearblock th
{
	padding-left: 0.33in;
	padding-right: 0.33in;
}

.yeartitle
{
	font-size: 90pt;
	letter-spacing: 0.5em;
	text-align: center;
	background-color: #000066;
	color: white;
	font-family: Arial Black, Arial, sans-serif;
	padding: 0;
	margin: 0;
}

.monthblock
{
	border-collapse: collapse;
	padding: 0;
}

.monthblock td
{
	padding: 0;
}

.monthtitle
{
	color: #003399;
	font-size: 40pt;
	font-family: Arial Black, Arial, sans-serif;
}

.daytitle
{
	font-size: 10pt;
	text-align: center;
	width: 14.28%;
	padding: 0;
}

.emptyday
{
	background-color: #EEEEEE;
	border: 1px solid #CCCCCC;
	border-left: 1px solid #000066;
	border-right: 1px solid #000066;
	width: 14.28%;
}

td.dayblock
{
	font-size: 48pt;
	width: 14.28%;
	font-weight: bold;
	font-family: Arial Narrow, Arial, sans-serif;
	text-align: center;
	vertical-align: middle;
	border: 1px solid #CCCCCC;
	border-left: 1px solid #000066;
	border-right: 1px solid #000066;
	padding-left: 0.08in;
	padding-right: 0.08in;
}

.reminders .remind_table td, .remind_date
{
	height: 6pt;
	font-size: 6pt;
	padding: 0;
	margin: 0;
	padding-right: 2px;
}

.weekend
{
	color: #999999;
}

.event
{
	/*color: #0033CC;*/
	text-decoration: underline;
}

.calendar_privatedef
{
	background-color: #CCEEFF;
}

.calendar_conference
{
	background-color: #FFCC99;
}

.calendar_holiday2
{
	background-color: #CCEEFF;
}

.calendar_holiday
{
	background-color: #AACCEE;
}

.calendar_usa
{
	color: #000099;
}

.calendar_season
{
	color: maroon;
}

.calendar_time
{
	color: purple;
}

.calendar_taxes
{
	color: green;
}

.calendar_family
{
	background-color: #FFCCCC;
}

.calendar_funday
{
	background-color: #FF9966;
}

.calendar_district
{
	color: red;
}

.calendar_rally
{
	color: #CC6600;
}

.calendar_al_rally
{
	background-color: #FFCC99;
	color: black;
}

.calendar_charity
{
	color: #CC66FF;
}

.calendar_safety
{
	background-color: #99FF99;
}




</style>
</head>

<body>

<%

FUNCTION loadRemindFiles( nDateBegin, nDateEnd, sCategories, sHolidayCategories, bShowToday )

	DIM	sRemindFile
	DIM	oCalendar
	DIM	oCalendarFile
	
	DIM	aFileList()
	REDIM aFileList(5)
	DIM	aTempSplit
	DIM	sTemp
	DIM	dateTemp
	
	getRemindList aFileList
	
	
	SET loadRemindFiles = Nothing
	
	IF -1 < UBOUND(aFileList) THEN
	
		aTempSplit = SPLIT( aFileList(0), vbTAB )
		
		SET	oCalendar = new CCalendar
		
		IF bShowToday THEN
			oCalendar.juliandate = jdateFromVBDate( NOW )
			oCalendar.subject = "* T * O * D * A * Y *"
			oCalendar.style = "RmdToday"
			oCalendar.outputMessage
		END IF

		DIM	sSavePictureFunc
		sSavePictureFunc = g_htmlFormat_pictureFunc
		g_htmlFormat_pictureFunc = "remindPicture"

		FOR EACH sTemp IN aFileList
			IF "" <> sTemp THEN
				aTempSplit = SPLIT( sTemp, vbTAB )
				IF 0 < INSTR(aTempSplit(1), ";key")  OR  0 < INSTR(LCASE(aTempSplit(1)), "motorcycle") THEN
					dateTemp = CDATE(aTempSplit(2))
					SET oCalendarFile = new CCalendarFile
					oCalendarFile.file = aTempSplit(1)
					oCalendarFile.datebegin = nDateBegin
					oCalendarFile.dateend = nDateEnd
					IF 0 < INSTR(aTempSplit(1),";categories") THEN
						IF 0 < LEN(sCategories) THEN oCalendarFile.categories = sCategories
					ELSEIF 0 < INSTR(aTempSplit(1),";holiday") THEN
						IF 0 < LEN(sHolidayCategories) THEN oCalendarFile.categories = sHolidayCategories
					END IF
					oCalendarFile.getDates( oCalendar )
				END IF
			END IF
		NEXT 'sTemp
				
		g_htmlFormat_pictureFunc = sSavePictureFunc
		SET loadRemindFiles = oCalendar
		SET oCalendar = Nothing
		SET oCalendarFile = Nothing

	END IF
END FUNCTION










DIM	gRows
DIM	gCols

SELECT CASE LCASE(q_layout)
CASE "portrait"
	gCols = 3
	gRows = 4
CASE "landscape"
	gCols = 4
	gRows = 3
CASE ELSE
	gCols = 0
	gCols = 0
END SELECT


FUNCTION getCategories( oXML, nD, nM, nY )

	getCategories = ""
	
	'EXIT FUNCTION
	
	DIM	oEvents
	DIM	oEvent
	DIM	oCategories
	DIM	oCategory
	
	DIM	sDate
	sDate = CSTR(nY) & " " & RIGHT("00"&nM,2) & " " & RIGHT("00"&nD,2)
	
	SET oEvents = oXML.documentElement.selectNodes( "/calendar/event[date = '" & sDate & "']")
	IF NOT Nothing IS oEvents THEN
		IF 0 < oEvents.length THEN
			getCategories = " event"
		END IF
	END IF
	
	FOR EACH oEvent IN oEvents
		SET oCategories = oEvent.selectNodes("category")
		FOR EACH oCategory IN oCategories
			getCategories = getCategories & " calendar_" & oCategory.text
		NEXT 'oCategory
	NEXT 'oEvent
	
	SET oCategory = Nothing
	SET oCategories = Nothing
	SET oEvent = Nothing
	SET oEvents = Nothing


END FUNCTION


SUB makeMonth( nM, nY )

	DIM	nWD
	DIM	i
	DIM	iw
	DIM	n
	DIM	bAbbrev
	
	DIM	sCategories
	
	IF "on" = LCASE(q_daynameabbrev) THEN
		bAbbrev = TRUE
	ELSE
		bAbbrev = FALSE
	END IF
	
	DIM	j
	DIM	jStart
	DIM	jEnd
	DIM	jW
	
	jStart = jdateFromGregorian( 1, nM, nY )
	jEnd = jdateFromGregorian( -1, nM, nY )
	jW = weekdayFromJDate( jStart )

	DIM	oXML
	DIM	oCalendar
	SET oCalendar = loadRemindFiles( jStart, jEnd, "newsletter", "none,holiday,holiday2", FALSE )
	IF NOT Nothing IS oCalendar THEN
	
		SET	oXML = oCalendar.xmldom
		
		'Response.Write Server.HTMLEncode( oCalendar.xml )
	
		DIM	oXSL
		SET oXSL = Server.CreateObject("msxml2.DOMDocument")
		'SET oXSL = Server.CreateObject("Microsoft.XMLDOM")
		oXSL.async = false
		oXSL.load(Server.MapPath("scripts/remind_yag.xslt"))
		
	END IF


	
	'Response.Write "jStart = " & jStart & "<br>"
	'Response.Write "jEnd = " & jEnd & "<br>"
	'Response.Write "jW = " & jW & "<br>"

	Response.Write "<table class=""monthblock"">" & vbCRLF
	
	Response.Write "<tr>" & vbCRLF
	IF 0 < CINT(q_startWeek) THEN
		nWD = CINT(q_startWeek)
		FOR i = 1 TO 7
			Response.Write "<td class=""daytitle"">" & WEEKDAYNAME( nWD, bAbbrev ) & "</td>" & vbCRLF
			nWD = nWD + 1
			IF 7 < nWD THEN nWD = 1
		NEXT 'i
	ELSE
	END IF
	Response.Write "</tr>" & vbCRLF
	
	nWD = CINT(q_startWeek)
	IF nWD <= jW THEN
		i = jW - nWD
	ELSE
		i = (jW+7) - nWD
	END IF
	nWD = 7 - i
	iw = 1
	Response.Write "<tr>" & vbCRLF
	DO WHILE 0 < i
		Response.Write "<td class=""emptyday""></td>" & vbCRLF
		i = i - 1
		iw = iw + 1
	LOOP
	j = jStart
	i = 1
	DO WHILE 0 < nWD
		sCategories = getCategories( oXML, i, nM, nY )
		Response.Write "<td class=""dayblock"
		SELECT CASE q_defWeekend
		CASE "fl"
			IF 7 = iw  OR  1 = iw THEN
				Response.Write " weekend"
			END IF
		CASE "f2"
			IF iw <=2 THEN
				Response.Write " weekend"
			END IF
		CASE "l2"
			IF 6 <= iw THEN
				Response.Write " weekend"
			END IF
		CASE ELSE
		END SELECT
		'IF 0 < INSTR(sCategories, ",holiday2") THEN
		'	Response.Write " holiday2"
		'ELSEIF 0 < INSTR(sCategories, ",holiday") THEN
		'	Response.Write " holiday"
		'ELSEIF 0 < INSTR(sCategories, ",ingr") THEN
		'	Response.Write " holiday2"
		'END IF
		IF "" <> sCategories THEN
			Response.Write sCategories
		END IF
		Response.Write """>" & i & "</td>" & vbCRLF
		i = i + 1
		j = j + 1
		iw = iw + 1
		IF 7 < iw THEN iw = 1
		nWD = nWD - 1
	LOOP
	Response.Write "</tr>" & vbCRLF
	
	Response.Write "<tr>" & vbCRLF
	iw = 1
	DO WHILE j <= jEnd
		sCategories = getCategories( oXML, i, nM, nY )
		Response.Write "<td class=""dayblock"
		SELECT CASE q_defWeekend
		CASE "fl"
			IF 7 = iw  OR  1 = iw THEN
				Response.Write " weekend"
			END IF
		CASE "f2"
			IF iw <=2 THEN
				Response.Write " weekend"
			END IF
		CASE "l2"
			IF 6 <= iw THEN
				Response.Write " weekend"
			END IF
		CASE ELSE
		END SELECT
		'IF 0 < INSTR(sCategories, ",holiday2") THEN
		'	Response.Write " holiday2"
		'ELSEIF 0 < INSTR(sCategories, ",holiday") THEN
		'	Response.Write " holiday"
		'ELSEIF 0 < INSTR(sCategories, ",ingr") THEN
		'	Response.Write " holiday2"
		'END IF
		IF "" <> sCategories THEN
			Response.Write sCategories
		END IF
		Response.Write """>" & i & "</td>" & vbCRLF
		i = i + 1
		j = j + 1
		iw = iw + 1
		IF 7 < iw THEN
			iw = 1
			Response.Write "</tr>" & vbCRLF
			Response.Write "<tr>" & vbCRLF
		END IF
	LOOP
	IF 1 < iw THEN
		DO WHILE iw <= 7
			Response.Write "<td class=""emptyday""></td>" & vbCRLF
			iw = iw + 1
		LOOP
	END IF
	Response.Write "</tr>" & vbCRLF
	
	Response.Write "</table>" & vbCRLF
	
	IF NOT Nothing IS oCalendar THEN

		Response.Write "<table class=""reminders"">" & vbCRLF
		Response.Write "<tr><td align=""left"">"
		Response.Write oXML.transformNode(oXSL)
		Response.Write "</td></tr>"
		Response.Write "</table>" & vbCRLF

	END IF
	
	
	SET oXML = Nothing
	SET oXSL = Nothing
	SET oCalendar = Nothing


END SUB


SUB progressionHorizontal()

	Response.Write "<div class=""yearpage"">"
	Response.Write "<div class=""yeartitle"">" & q_year
	IF 1 <> q_startMonth THEN
		Response.Write " / " & CSTR(q_year+1)
	END IF
	Response.Write "</div>" & vbCRLF
	
	Response.Write "<table class=""yearblock"">" & vbCRLF
	
	DIM	nR
	DIM	nC
	DIM	nM
	DIM	nY
	
	nM = q_startMonth
	nY = q_year
	FOR nR = 1 TO gRows
		Response.Write "<tr>" & vbCRLF
		
		FOR nC = 1 TO gCols
			Response.Write "<td valign=""top"" width=""" & INT(100/gCols) & "%"" class=""monthcell"">"
			
			Response.Write "<div class=""monthtitle"">"
			Response.Write MONTHNAME( nM, FALSE )
			Response.Write "</div>" & vbCRLF
			
			makeMonth nM, nY
			
			nM = nM + 1
			IF 12 < nM THEN
				nM = 1
				nY = nY + 1
			END IF
			
			Response.Write "</td>" & vbCRLF
		NEXT 'nC
		
		Response.Write "</tr>" & vbCRLF
	NEXT 'nR
	
	Response.Write "</table>" & vbCRLF
	Response.Write "</div>" & vbCRLF

END SUB



SUB progressionVertical()

	Response.Write "<div class=""yearpage"">"
	Response.Write "<div class=""yeartitle"">" & q_year
	IF 1 <> q_startMonth THEN
		Response.Write " / " & CSTR(q_year+1)
	END IF
	Response.Write "</div>" & vbCRLF
	
	Response.Write "<table class=""yearblock"">" & vbCRLF
	
	DIM	nR
	DIM	nC
	DIM	nM
	DIM	nMx
	DIM	nY
	DIM	nYx
	
	nM = q_startMonth
	nMx = nM
	nY = q_year
	nYx = nY
	FOR nR = 1 TO gRows
		Response.Write "<tr>" & vbCRLF
		
		nM = nMx
		nY = nYx
		
		FOR nC = 1 TO gCols
			Response.Write "<td valign=""top"" width=""" & INT(100/gCols) & "%"" class=""monthcell"">"
			
			Response.Write "<div class=""monthtitle"">"
			Response.Write MONTHNAME( nM, FALSE )
			Response.Write "</div>" & vbCRLF
			
			makeMonth nM, nY
			
			nM = nM + gRows
			IF 12 < nM THEN
				nM = nM MOD 12
				nY = nY + 1
			END IF
			
			Response.Write "</td>" & vbCRLF
		NEXT 'nC
		
		nMx = nMx + 1
		IF 12 < nMx THEN
			nMx = nMx MOD 12
			nYx = nYx + 1
		END IF
		
		Response.Write "</tr>" & vbCRLF
	NEXT 'nR
	
	Response.Write "</table>" & vbCRLF
	Response.Write "</div>" & vbCRLF

END SUB



SELECT CASE LCASE(q_Progression)
CASE "horizontal"
	progressionHorizontal
CASE "vertical"
	progressionVertical
CASE ELSE
END SELECT



%>

</body>

</html>
<%

SET g_oFSO = Nothing
%>