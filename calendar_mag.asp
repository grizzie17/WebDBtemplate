<%@ Language=VBScript %>
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




DIM	q_categories
q_categories = Request("cat")

DIM q_m
q_m = Request("m")
IF "" = q_m THEN
	q_m = "1"
END IF

DIM	nMonths
nMonths = CINT(q_m)


IF "localhost" = LCASE(Request.ServerVariables("HTTP_HOST")) THEN
	IF "" = q_categories THEN
	'	q_categories = "~announcement:~marker"
	ELSE
	'	q_categories = q_categories & ":~announcement:~marker"
	END IF
END IF



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

q_year = YEAR(NOW)
q_startWeek = 2		' Monday -- put weekend as the end of the week
q_daynameabbrev = "on"
q_startMonth = MONTH(NOW)
q_layout = "Portrait"
q_Progression = "Horizontal"
q_defWeekend = "l2"


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts\remind.asp"-->
<!--#include file="scripts\remind_files.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
Response.Write "<title>" & Server.HTMLEncode("Month at a Glance - " & g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Month at a Glance - Calendar</title>
<%
END IF
%>
<!--#include file="scripts\favicon.asp"-->
<link rel="stylesheet" href="<%=remindCSS()%>" type="text/css">
<link rel="stylesheet" href="calendar_mag.css" type="text/css">
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>

</head>

<body>

<%










DIM	gRows
DIM	gCols

SELECT CASE LCASE(q_layout)
CASE "portrait"
	gCols = 1
	gRows = 12
CASE "landscape"
	gCols = 4
	gRows = 3
CASE ELSE
	gCols = 0
	gCols = 0
END SELECT


FUNCTION getCategories( BYREF sTable, oXML, nD, nM, nY )

	getCategories = ""
	sTable = ""
	
	'EXIT FUNCTION
	
	DIM	oEvents
	DIM	oEvent
	DIM	oCategories
	DIM	oCategory
	DIM	oStyle
	DIM	oSubject
	DIM	oBodyText
	DIM	oLocation
	DIM	oComments
	
	DIM	sStyle
	
	DIM	sDate
	sDate = CSTR(nY) & " " & RIGHT("00"&nM,2) & " " & RIGHT("00"&nD,2)
	
	SET oEvents = oXML.selectNodes( "/calendar/event[date = '" & sDate & "']")
	IF NOT Nothing IS oEvents THEN
		IF 0 < oEvents.length THEN
			getCategories = " event"
			sTable = "<table border=""0"" class=""reminders"">"
			
			FOR EACH oEvent IN oEvents
				SET oSubject = oEvent.selectSingleNode("subject")
				IF NOT Nothing IS oSubject THEN
					sTable = sTable & "<tr>"
					SET oStyle = oEvent.selectSingleNode("style")
					IF Nothing IS oStyle THEN
						sStyle = "normal"
					ELSE
						sStyle = oStyle.text
					END IF
					sTable = sTable & "<td class=""" & sStyle & """>-</td>"
					sTable = sTable & "<td class=""" & sStyle & """>"
					sTable = sTable & oSubject.text
					SET oBodyText = oEvent.selectSingleNode("content")
					IF NOT Nothing IS oBodyText THEN
						sTable = sTable & " -- <i>" & oBodyText.text & "</i>"
					END IF
					SET oLocation = oEvent.selectSingleNode("location")
					IF NOT Nothing IS oLocation THEN
						sTable = sTable & " (" & oLocation.text & ")"
					END IF
					SET oComments = oEvent.selectSingleNode("comments")
					IF NOT Nothing IS oComments THEN
						sTable = sTable & "<br>" & oComments.text
					END IF
					sTable = sTable & "</td></tr>"
					SET oCategories = oEvent.selectNodes("category")
					FOR EACH oCategory IN oCategories
						getCategories = getCategories & " calendar_" & oCategory.text
					NEXT 'oCategory
				END IF
			NEXT 'oEvent
			
			sTable = sTable & "</table>"
		END IF
	END IF
	
	
	SET oCategory = Nothing
	SET oCategories = Nothing
	SET oEvent = Nothing
	SET oEvents = Nothing


END FUNCTION


FUNCTION dummyPictureGen( sAlign, sBorder, sWidth, sColor, sCaption, sFile )
	dummyPictureGen = " "
END FUNCTION


FUNCTION dummyPictureFile( sLabel )
	dummyPictureFile = " "
END FUNCTION


SUB makeMonth( nM, nY )

	DIM	nWD
	DIM	i
	DIM	iw
	DIM	n
	DIM	bAbbrev
	DIM	sTable
	
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


	DIM	sSavePictureFunc
	sSavePictureFunc = g_htmlFormat_pictureFunc
	g_htmlFormat_pictureFunc = "dummyPictureFile"
	g_htmlFormat_pictureGenFunc = "dummyPictureGen"


	DIM	oXML
	DIM	oXMLc
	DIM	oCalendar
	SET oCalendar = loadRemindFiles2( jStart, jEnd, _
				q_categories, "holiday,usa,none", FALSE, NOW, "mag" )
	IF NOT Nothing IS oCalendar THEN
	
		SET	oXMLc = oCalendar.xmldom
		
		'Response.Write Server.HTMLEncode( oCalendar.xml )
	
		DIM	oXSL
		SET oXSL = remindLoadXmlFile( Server.MapPath("scripts/remind_mag.xslt") )
		'SET oXSL = Server.CreateObject("msxml2.DOMDocument")
		SET oXML = Server.CreateObject("msxml2.DOMDocument.6.0")
		'SET oXSL = Server.CreateObject("Microsoft.XMLDOM")
		'oXSL.async = false
		'oXSL.load(Server.MapPath("scripts/remind_mag.xslt"))
		
		oXMLc.transformNodeToObject oXSL, oXML
		
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
		sCategories = getCategories( sTable, oXML, i, nM, nY )
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
		Response.Write """>"
		Response.Write "<div class=""daynumber"">" & i & "</div>"
		IF "" <> sTable THEN
			Response.Write sTable
		ELSE
			Response.Write "<p>&nbsp;&nbsp;</p>"
		END IF
		Response.Write "</td>" & vbCRLF
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
		sCategories = getCategories( sTable, oXML, i, nM, nY )
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
		Response.Write """>"
		Response.Write "<div class=""daynumber"">" & i & "</div>"
		IF "" <> sTable THEN
			Response.Write sTable
		ELSE
			Response.Write "<p>&nbsp;&nbsp;</p>"
		END IF
		Response.Write "</td>" & vbCRLF
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
	
	
	
	SET oXML = Nothing
	SET oXSL = Nothing
	SET oCalendar = Nothing


END SUB


SUB progressionHorizontal()

	Response.Write "<div class=""yearpage"">"
	'Response.Write "<div class=""yeartitle"">" & q_year
	'IF 1 <> q_startMonth THEN
	'	Response.Write " / " & CSTR(q_year+1)
	'END IF
	'Response.Write "</div>" & vbCRLF
	
	'Response.Write "<table class=""yearblock"">" & vbCRLF
	
	DIM	nR
	DIM	nC
	DIM	nM
	DIM	nY
	
	nM = q_startMonth
	nY = q_year
	FOR nR = 1 TO nMonths
		'Response.Write "<tr>" & vbCRLF
		
		'FOR nC = 1 TO gCols
			'Response.Write "<td valign=""top"" width=""" & INT(100/gCols) & "%"" class=""monthcell"">"
			
			Response.Write "<div class=""monthtitle"">"
			Response.Write MONTHNAME( nM, FALSE ) & " " & nY
			Response.Write "</div>" & vbCRLF
			
			makeMonth nM, nY
			
			nM = nM + 1
			IF 12 < nM THEN
				nM = 1
				nY = nY + 1
			END IF
			
			'Response.Write "</td>" & vbCRLF
		'NEXT 'nC
		
		'Response.Write "</tr>" & vbCRLF
	NEXT 'nR
	
	'Response.Write "</table>" & vbCRLF
	Response.Write "</div>" & vbCRLF

END SUB



SUB progressionVertical()

	Response.Write "<div class=""yearpage"">"
	'Response.Write "<div class=""yeartitle"">" & q_year
	'IF 1 <> q_startMonth THEN
	'	Response.Write " / " & CSTR(q_year+1)
	'END IF
	'Response.Write "</div>" & vbCRLF
	
	'Response.Write "<table class=""yearblock"">" & vbCRLF
	
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
		'Response.Write "<tr>" & vbCRLF
		
		nM = nMx
		nY = nYx
		
		FOR nC = 1 TO gCols
			'Response.Write "<td valign=""top"" width=""" & INT(100/gCols) & "%"" class=""monthcell"">"
			
			Response.Write "<div class=""monthtitle"">"
			Response.Write MONTHNAME( nM, FALSE ) & " " & nY
			Response.Write "</div>" & vbCRLF
			
			makeMonth nM, nY
			
			nM = nM + gRows
			IF 12 < nM THEN
				nM = nM MOD 12
				nY = nY + 1
			END IF
			
			'Response.Write "</td>" & vbCRLF
		NEXT 'nC
		
		nMx = nMx + 1
		IF 12 < nMx THEN
			nMx = nMx MOD 12
			nYx = nYx + 1
		END IF
		
		'Response.Write "</tr>" & vbCRLF
	NEXT 'nR
	
	'Response.Write "</table>" & vbCRLF
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
<!--#include file="scripts\include_db_end.asp"-->
