<%@	Language=VBScript
	EnableSessionState=True %>
<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2008 Bear Consulting Group
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
<!--#include file="scripts\dateutils.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<%




'===========================================================
DIM	oXML
DIM	sFilename
DIM	sID
DIM	oEvent
DIM	sDate
DIM	sDateType
DIM	sDateSubType
DIM	sDateOffset
DIM	sDateDuration
DIM	sDateDurationView
DIM	sDatePending
DIM	sSubject
DIM	sContent
DIM sLocation
DIM sComments
DIM	sCategories
DIM	sStyle
DIM	sTime
PUBLIC aDateParse

DIM	sRemindFile

SET oEvent = Nothing
sDate = ""
sDateType = ""
sDateSubType = ""
sDateOffset = ""
sDateDuration = ""
sDateDurationView = ""
sDatePending = ""
sSubject = ""
sContent = ""
sLocation = ""
sComments = ""
sCategories = ""
sStyle = ""
sTime = ""

sFilename = Request("filename")
IF 0 = LEN(sFilename) THEN sFilename = "remind.xml"

sID = Request.QueryString("id")
IF 0 < LEN(sID) THEN
	sRemindFile = findRemindFile( sFilename )
	SET oXML = remindLoadXmlFile( sRemindFile )
	SET oEvent = oXML.selectSingleNode("//event[@id=""" & sID & """]")
	IF Nothing IS oEvent THEN
		sSubject = "unable to resolve ID"
	ELSE
		DIM	oDate
		DIM	oTemp
		DIM	oTempList
		SET oDate = oEvent.selectSingleNode("date")
		sDate = oDate.text
		aDateParse = SPLIT(sDate," ",-1,vbTextCompare)
		SET oTemp = oDate.attributes.getNamedItem("type")
		IF NOT Nothing IS oTemp THEN sDateType = oTemp.value
		SELECT CASE sDateType
		CASE "single"
		CASE "monthly"
			SET oTemp = oDate.attributes.getNamedItem("monthly")
			sDateSubType = oTemp.value
		CASE "yearly"
			SET oTemp = oDate.attributes.getNamedItem("yearly")
			sDateSubType = oTemp.value
		CASE "hebrew"
			SET oTemp = oDate.attributes.getNamedItem("hebrew")
			sDateSubType = oTemp.value
		CASE "keyword"
		END SELECT
		SET oTemp = oDate.attributes.getNamedItem("time")
		IF NOT Nothing IS oTemp THEN sTime = TRIM(oTemp.value)
		SET oTemp = oDate.attributes.getNamedItem("duration")
		IF NOT Nothing IS oTemp THEN sDateDuration = TRIM(oTemp.value)
		SET oTemp = oDate.attributes.getNamedItem("duration-view")
		IF NOT Nothing IS oTemp THEN sDateDurationView = TRIM(oTemp.value)
		IF "" = sDateDurationView THEN sDateDurationView = "all"
		SET oTemp = oDate.attributes.getNamedItem("offset")
		IF NOT Nothing IS oTemp THEN sDateOffset = TRIM(oTemp.value)
		SET oTemp = oDate.attributes.getNamedItem("pending")
		IF NOT Nothing IS oTemp THEN sDatePending = TRIM(oTemp.value)
		
		sCategories = ""
		SET oTempList = oEvent.getElementsByTagName("category")
		FOR EACH oTemp IN oTempList
			IF 0 < LEN(sCategories) THEN
				sCategories = sCategories & ", " & oTemp.text
			ELSE
				sCategories = oTemp.text
			END IF
		NEXT 'oTemp

		SET oTemp = oEvent.selectSingleNode("style")
		IF NOT Nothing IS oTemp THEN sStyle = oTemp.text
		
		SET oTemp = oEvent.selectSingleNode("subject")
		sSubject = oTemp.text
		SET oTemp = oEvent.selectSingleNode("content")
		IF NOT Nothing IS oTemp THEN sContent = TRIM(oTemp.text)
		
		SET oTemp = oEvent.selectSingleNode("location")
		IF NOT Nothing IS oTemp THEN sLocation = TRIM(oTemp.text)
		
		SET oTemp = oEvent.selectSingleNode("comments")
		IF NOT Nothing IS oTemp THEN sComments = TRIM(oTemp.text)
		
	END IF
	SET oTemp = Nothing
	SET oEvent = Nothing
	SET oXML = Nothing
ELSE
	sDateType = "single"
	sDate = YEAR(NOW) & " " & MONTH(NOW) & " " & DAY(NOW)
	aDateParse = SPLIT(sDate," ",-1,vbTextCompare)
END IF



Function WordCase( s )

    Dim aList
    aList = Split(s, " ")
    
    Dim i
    For i = LBound(aList) To UBound(aList)
        Select Case UCase(aList(i))
        Case "US", "PO", "AL", "NE", "NW", "SE", "SW"
            aList(i) = UCase(aList(i))
        Case Else
            If "MC" = Left(aList(i), 2) Then
                aList(i) = UCase(Left(aList(i), 1)) & LCase(Mid(aList(i), 2, 1)) & UCase(Mid(aList(i), 3, 1)) & LCase(Mid(aList(i), 4))
            Else
                aList(i) = UCase(Left(aList(i), 1)) & LCase(Mid(aList(i), 2))
            End If
        End Select
    Next 'i
    
    WordCase = Join(aList, " ")

End Function



SUB getCSStyles()
	DIM	oFile
	DIM	sFile
	DIM	sText
	DIM	sTemp
	
	sFile = findRemindFile( "remind.css" )
	SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
	sText = oFile.ReadAll()
	oFile.Close
	SET oFile = Nothing
	
	DIM	oReg
	DIM	oMatchList
	DIM	oMatch
	
	DIM	sClass
	DIM	sComment
	
	SET oReg = NEW RegExp
	'oReg.Pattern = "\s\.(\w+)\s*\{"
	oReg.Pattern = "\s\.(\w+)\s*\{\s*(/\*\s*::\s*([^:]+)(::|\*/))?"
	oReg.IgnoreCase = TRUE
	oReg.Global = TRUE
	SET oMatchList = oReg.Execute( sText )
	FOR EACH oMatch IN oMatchList
		sClass = oMatch.Submatches(0)
		IF "RmdToday" <> sClass THEN
			IF 1 < oMatch.Submatches.Count THEN
				sComment = TRIM(oMatch.Submatches(1))
			ELSE
				sComment = ""
			END IF
			Response.Write "<option value=""" & sClass & """ class=""" & sClass & """>"
			IF "" = sComment THEN
				sTemp = REPLACE(sClass,"Rmd","")
				sTemp = REPLACE( sTemp, "_", " " )
				sTemp = WordCase( sTemp )
				Response.Write sTemp
			ELSE
				sTemp = REPLACE(sComment,"/*","")
				sTemp = REPLACE(sTemp,"::","")
				sTemp = TRIM(sTEMP)
				Response.Write Server.HTMLEncode(sTemp)
			END IF
			Response.Write "</option>" & vbCRLF
		END IF
	NEXT 'oMatch
	SET oMatch = Nothing
	SET oMatchList = Nothing
	SET oReg = Nothing

END SUB


FUNCTION checkType( sType )
	checkType = ""
	IF sDateType = sType THEN
		checkType = "checked"
	END IF
END FUNCTION
FUNCTION checkYearlyType( sType )
	checkYearlyType = ""
	IF sDateSubType = sType THEN
		checkYearlyType = "checked"
	END IF
END FUNCTION
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="<%=remindCSS()%>">
<script type="text/javascript" language="javascript" src="remind_edit.js"></script>
<title>Edit Calendar Event</title>
</head>

<body>

<form method="POST" name="formDate" action="remind_editsubmit.asp" onsubmit="submitValidate(this)">
	<%
IF 0 < LEN(sFilename) THEN
	Response.Write "<input type=""hidden"" name=""filename"" value=""" & sFilename & """>" & vbCRLF
END IF
IF 0 < LEN(sID) THEN
	Response.Write "<input type=""hidden"" name=""id"" value=""" & sID & """>" & vbCRLF
END IF
%>
	<table border="0" width="100%" cellspacing="0" cellpadding="4">
		<tr>
			<th bgcolor="#CCFFCC" align="left">Edit Event / Message</th>
			<td align="right" bgcolor="#CCFFCC">
			<a href="javascript:window.history.back()">Cancel (Back)</a>&nbsp;
			<input type="submit" value="Save" name="Save"></td>
		</tr>
	</table>
	<table border="0" cellspacing="1">
		<tr>
			<td align="right" valign="top">Subject:</td>
			<td>
			<input type="text" name="subject" onblur="fixupContent(this)" size="60" value="<%=Server.HTMLEncode(sSubject)%>"></td>
		</tr>
		<tr>
			<td align="right" valign="top">Body</td>
			<td><textarea rows="3" name="content" onblur="fixupContent(this)" cols="50"><%=Server.HTMLEncode(sContent)%></textarea></td>
		</tr>
		<tr>
			<td align="right" valign="top">Location</td>
			<td><textarea rows="3" name="location" onblur="fixupContent(this)" cols="50"><%=Server.HTMLEncode(sLocation)%></textarea></td>
		</tr>
		<tr>
			<td align="right" valign="top">Comments</td>
			<td><textarea rows="3" name="comments" onblur="fixupContent(this)" cols="50"><%=Server.HTMLEncode(sComments)%></textarea></td>
		</tr>
		<tr>
			<td align="right" valign="top">Style</td>
			<td valign="top">
			<!--input type="text" name="style" size="20" value="<%=sStyle%>"-->
			<select size="1" name="style">
			<option value>--default--</option>
			<%
		getCSStyles
%>
			</select></td>
		</tr>
		<tr>
			<td align="right" valign="top">Categories</td>
			<td valign="top">
			<input type="text" name="categories" size="50" value="<%=sCategories%>"><br>
			<font size="2">(comma separated)</font></td>
		</tr>
	</table>
	<table border="0" cellpadding="2">
		<tr>
			<th id="type_single" align="left">Date</th>
			<th id="content_single" align="left"></th>
		</tr>
		<tr>
			<td bgcolor="#CCCCCC" id="type_single">
			<input type="radio" value="single" <%=checktype("single")%> name="recurrence" id="datetypeSingle"><label for="datetypeSingle">Single</label></td>
			<td bgcolor="#CCCCCC" id="content_single">
			<input type="text" name="single_day" size="3" value="1" onchange="onDay(this)"> 
			-
			<select size="1" name="single_month" onchange="onSelectMonth(this)">
			<option value="1">January</option>
			<option value="2">February</option>
			<option value="3">March</option>
			<option value="4">April</option>
			<option value="5">May</option>
			<option value="6">June</option>
			<option value="7">July</option>
			<option value="8">August</option>
			<option value="9">September</option>
			<option value="10">October</option>
			<option value="11">November</option>
			<option value="12">December</option>
			</select> -
			<input type="text" name="single_year" size="5" value="<%=YEAR(NOW)%>">
			</td>
		</tr>
		<!--tr>
      <td>
        <input type="radio" value="daily" name="recurrence"
        id="datetypeDaily"><label
        for="datetypeDaily">Daily</label></td>
      <td>
            <input type="radio" name="radioDaily" value="count" checked id="fp5">Every
            <input type="text" name="daily_days" size="3" value="1">  day(s)<br>
            <input type="radio" name="radioDaily" value="weekday" id="fp6">Every
            weekday</td>
    </tr-->
		<tr>
      <td bgcolor="#EEEEEE">
        <input type="radio" value="weekly" <%=checktype("weekly")%> name="recurrence"
        id="datetypeWeekly"><label
        for="datetypeWeekly">Weekly</label></td>
      <td bgcolor="#EEEEEE">Recur every <select size="1" name="weeklyRepeat">
      <option value="0">Every Week</option>
      <option value="1">Odd Weeks</option>
      <option value="2">Even Weeks</option>
      </select>
        on:
            <table border="0" cellspacing="0" width="100%">
              <tr>
                <td><input type="checkbox" name="weeklySunday" value="ON"
                  id="fp10"><label for="fp10"> Sunday</label></td>
                <td><input type="checkbox" name="weeklyMonday" value="ON"
                  id="fp11"><label for="fp11"> Monday</label></td>
                <td><input type="checkbox" name="weeklyTuesday" value="ON"
                  id="fp12"><label for="fp12"> Tuesday</label></td>
                <td><input type="checkbox" name="weeklyWednesday" value="ON"
                  id="fp13"><label for="fp13"> Wednesday</label></td>
              </tr>
              <tr>
                <td><input type="checkbox" name="weeklyThursday" value="ON"
                  id="fp9"><label for="fp9"> Thursday</label></td>
                <td><input type="checkbox" name="weeklyFriday" value="ON"
                  id="fp8"><label for="fp8"> Friday</label></td>
                <td><input type="checkbox" name="weeklySaturday" value="ON"
                  id="fp7"><label for="fp7"> Saturday</label></td>
                <td></td>
              </tr>
            </table>
      </td>
    </tr>
		<tr>
			<td rowspan="2" bgcolor="#CCCCCC">
			<input type="radio" value="monthly" <%=checktype("monthly")%> name="recurrence" id="datetypeMonthly"><label for="datetypeMonthly">Monthly</label></td>
			<td bgcolor="#CCCCCC">
			<input type="radio" name="radioMonthly" value="dayn" <%=checkyearlytype("dayn")%> checked id="fp16"><label for="fp16">Day</label>
			<input type="text" name="monthlyDayn_day" size="3" value="1"> of
			<select size="1" name="monthlyDayn_cycle" onchange="onSelectCycle(this)">
			<option value="1">Every Month</option>
			<option value="2">Bimonthly</option>
			<option value="3">Quarterly</option>
			<option value="6">Semiannually</option>
			</select> cycle
			<select size="1" name="monthlyDayn_index" onchange="onSelectMonthIndex(this)">
			<option>1</option>
			<option>2</option>
			<option>3</option>
			<option>4</option>
			<option>5</option>
			<option>6</option>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#CCCCCC">
			<input type="radio" name="radioMonthly" value="wday" <%=checkyearlytype("wday")%> id="fp17"><label for="fp17">The</label>
			<select size="1" name="monthlyWDay_weekNumber" onchange="onWeeknumber(this)">
			<option value="1">first</option>
			<option value="2">second</option>
			<option value="3">third</option>
			<option value="4">fourth</option>
			<option value="5">fifth</option>
			<option value="0">-----</option>
			<option value="-3">3rd from last</option>
			<option value="-2">next to last</option>
			<option value="-1">last</option>
			</select><select size="1" name="monthlyWDay_weekday" onchange="onSelectWeekday(this)">
			<option value="1">Sunday</option>
			<option value="2">Monday</option>
			<option value="3">Tuesday</option>
			<option value="4">Wednesday</option>
			<option value="5">Thursday</option>
			<option value="6">Friday</option>
			<option value="7">Saturday</option>
			<option value="-">-----</option>
			<option value="-1">weekday</option>
			<option value="-2">weekend</option>
			</select> of&nbsp;
			<select size="1" name="monthlyWDay_cycle" onchange="onSelectCycle(this)">
			<option value="1">Every Month</option>
			<option value="2">Bimonthly</option>
			<option value="3">Quarterly</option>
			<option value="6">Semiannually</option>
			</select> cycle
			<select size="1" name="monthlyWDay_index" onchange="onSelectMonthIndex(this)">
			<option>1</option>
			<option>2</option>
			<option>3</option>
			<option>4</option>
			<option>5</option>
			<option>6</option>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#EEEEEE" rowspan="2" id="type_yearly">
			<input type="radio" value="yearly" <%=checktype("yearly")%> name="recurrence" id="datetypeYearly"><label for="datetypeYearly">Yearly</label></td>
			<td bgcolor="#EEEEEE" id="content_yearlyDayn">
			<input type="radio" name="radioYearly" value="dayn" <%=checkyearlytype("dayn")%> checked id="fp15"><label for="fp15"> 
			Every</label>
			<select size="1" name="yearlyDayn_month" onchange="onSelectMonth(this)">
			<option value="1">January</option>
			<option value="2">February</option>
			<option value="3">March</option>
			<option value="4">April</option>
			<option value="5">May</option>
			<option value="6">June</option>
			<option value="7">July</option>
			<option value="8">August</option>
			<option value="9">September</option>
			<option value="10">October</option>
			<option value="11">November</option>
			<option value="12">December</option>
			</select><input type="text" name="yearlyDayn_day" size="3" value="1" onchange="onDay(this)">
			</td>
		</tr>
		<tr>
			<td bgcolor="#EEEEEE" id="content_yearlyWDay">
			<input type="radio" name="radioYearly" value="wday" <%=checkyearlytype("wday")%> id="fp14"><label for="fp14"> 
			The</label>
			<select size="1" name="yearlyWDay_weekNumber" onchange="onWeeknumber(this)">
			<option value="1">first</option>
			<option value="2">second</option>
			<option value="3">third</option>
			<option value="4">fourth</option>
			<option value="5">fifth</option>
			<option value="0">-----</option>
			<option value="-3">3rd from last</option>
			<option value="-2">next to last</option>
			<option value="-1">last</option>
			</select><select size="1" name="yearlyWDay_weekday" onchange="onSelectWeekday(this)">
			<option value="1">Sunday</option>
			<option value="2">Monday</option>
			<option value="3">Tuesday</option>
			<option value="4">Wednesday</option>
			<option value="5">Thursday</option>
			<option value="6">Friday</option>
			<option value="7">Saturday</option>
			<option value="-">-----</option>
			<option value="-1">weekday</option>
			<option value="-2">weekend</option>
			</select> of
			<select size="1" name="yearlyWDay_month" onchange="onSelectMonth(this)">
			<option value="1">January</option>
			<option value="2">February</option>
			<option value="3">March</option>
			<option value="4">April</option>
			<option value="5">May</option>
			<option value="6">June</option>
			<option value="7">July</option>
			<option value="8">August</option>
			<option value="9">September</option>
			<option value="10">October</option>
			<option value="11">November</option>
			<option value="12">December</option>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#CCCCCC" rowspan="2" id="type_hebrew">
			<input type="radio" value="hebrew" <%=checktype("hebrew")%> name="recurrence" id="datetypeHebrew"><label for="datetypeHebrew">Hebrew</label></td>
			<td bgcolor="#CCCCCC" id="content_hebrewDayn">
			<input type="radio" name="radioHebrew" value="dayn" <%=checkyearlytype("dayn")%> checked id="fp15"><label for="fp15"> 
			Every</label>
			<select size="1" name="hebrewDayn_month" onchange="onSelectHebrewMonth(this)">
			<option value="1">Nisan</option>
			<option value="2">Iyyar</option>
			<option value="3">Sivan</option>
			<option value="4">Tammuz</option>
			<option value="5">Av</option>
			<option value="6">Elul</option>
			<option value="7">Tishri</option>
			<option value="8">Heshvan</option>
			<option value="9">Kislev</option>
			<option value="10">Teveth</option>
			<option value="11">Shevat</option>
			<option value="12">Adar</option>
			<option value="13">Veadar</option>
			</select><input type="text" name="hebrewDayn_day" size="3" value="1" onchange="onDay(this)">
			</td>
		</tr>
		<tr>
			<td bgcolor="#CCCCCC" id="content_HebrewWDay">
			<input type="radio" name="radioHebrew" value="wday" <%=checkyearlytype("wday")%> id="fp14"><label for="fp14"> 
			The</label>
			<select size="1" name="hebrewWDay_weekNumber" onchange="onWeeknumber(this)">
			<option value="1">first</option>
			<option value="2">second</option>
			<option value="3">third</option>
			<option value="4">fourth</option>
			<option value="5">fifth</option>
			<option value="0">-----</option>
			<option value="-3">3rd from last</option>
			<option value="-2">next to last</option>
			<option value="-1">last</option>
			</select><select size="1" name="hebrewWDay_weekday" onchange="onSelectWeekday(this)">
			<option value="1">Sunday</option>
			<option value="2">Monday</option>
			<option value="3">Tuesday</option>
			<option value="4">Wednesday</option>
			<option value="5">Thursday</option>
			<option value="6">Friday</option>
			<option value="7">Saturday</option>
			<option value="-">-----</option>
			<option value="-1">weekday</option>
			<option value="-2">weekend</option>
			</select> of
			<select size="1" name="hebrewWDay_month" onchange="onSelectHebrewMonth(this)">
			<option value="1">Nisan</option>
			<option value="2">Iyyar</option>
			<option value="3">Sivan</option>
			<option value="4">Tammuz</option>
			<option value="5">Av</option>
			<option value="6">Elul</option>
			<option value="7">Tishri</option>
			<option value="8">Heshvan</option>
			<option value="9">Kislev</option>
			<option value="10">Teveth</option>
			<option value="11">Shevat</option>
			<option value="12">Adar</option>
			<option value="13">Veadar</option>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#EEEEEE" id="type_keyword">
			<input type="radio" value="keyword" <%=checktype("keyword")%> name="recurrence" id="datetypeKeyword"><label for="datetypeKeyword">Keyword</label>
			</td>
			<td bgcolor="#EEEEEE" id="content_keyword">
			<select size="1" name="keyword_name">
			<option value="easter">Easter</option>
			<option value="orthodox easter">Orthodox Easter</option>
			<option value="paschal">Paschal</option>
			<option value="passover">Passover</option>
			<option value="rosh hashanah">Rosh Hashanah</option>
			<option value="moon full">Full Moon</option>
              <option value="moon new">New Moon</option></select> </td>
		</tr>
		<tr>
			<td bgcolor="#CCCCCC" id="type_season">
			<input type="radio" value="season" <%=checktype("season")%> name="recurrence" id="datetypeSeason"><label for="datetypeSeason">Season</label>
			</td>
			<td bgcolor="#CCCCCC" id="content_season">
			<select size="1" name="season_name">
			<option value="0">Equinox (March)</option>
			<option value="1">Solstice (June)</option>
			<option value="2">Equinox (September)</option>
			<option value="3">Solstice (December)</option>
			</select> </td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>Time: <input type="text" name="time" size="8" value="<%=sTime%>"> 24 hour time 
			(hh:mm)<br>
			Offset:
			<input type="text" name="offset" size="30" value="<%=sDateOffset%>">&nbsp; 
			<br>
			Duration:
			<input type="text" name="duration" size="5" value="<%=sDateDuration%>">&nbsp; Expand Duration View:<select name="durationview">
			<option value="all">All Views</option>
			<option value="list">List View</option>
			<option value="mag">Month at a Glance</option>
			<option value="none">No Views</option>
			</select><br>Pending:
			<input type="text" name="pending" size="5" value="<%=sDatePending%>"></td>
		</tr>
	</table>
	<table border="0" width="100%" bgcolor="#CCFFCC" cellspacing="0" cellpadding="4">
		<tr>
			<td><input type="reset" value="Reset" name="B2">&nbsp;&nbsp; </td>
			<td align="right">&nbsp;<input type="submit" value="Save" name="Save1">
			<input type="button" value="Cancel" name="Cancel" onclick="window.history.back()"></td>
		</tr>
	</table>
</form>
<script type="text/javascript" language="JavaScript">
<!--
var	oForm;
var n;

oForm = document.formDate;

function selectCurrentStyle()
{
	var	oStyle = oForm.style;
	var	sStyle = "<%=sStyle%>";
	if ( ! selectItemOnValue( oStyle, sStyle ) )
	{
		var oOption;
		var sName;
		
		sName = sStyle;
		if ( "Rmd" == sName.substr(0,3) )
			sName = sName.substr(3);
		
		oOption = document.createElement("OPTION");
		oStyle.options.add( oOption );
		oOption.text = sName;
		oOption.value = sStyle;
		oOption.style.color = "#99FF66";
		
		selectItemOnValue( oStyle, sStyle );
	}
}
selectCurrentStyle();

selectItemOnValue( oForm.durationview, "<%=sDateDurationView%>" );

<%
DIM	dd, mm, yy
DIM	wn, wd
DIM	jd
SELECT CASE sDateType
CASE "single"
	%>
	oForm.single_day.value = <%=aDateParse(2)%>;
	onDay( oForm.single_day );
	n = <%=aDateParse(1)%>;
	selectItemOnValue( oForm.single_month, n );
	onSelectMonth( oForm.single_month );
	oForm.single_year.value = <%=aDateParse(0)%>;
	<%
	yy = CINT(aDateParse(0))
	mm = CINT(aDateParse(1))
	dd = CINT(aDateParse(2))
	jd = jdateFromGregorian( dd, mm, yy )
	wd = weekdayFromJDate( jd )
	wn = (dd - 1) \ 7 + 1
	%>
	selectItemOnValue( oForm.yearlyWDay_weekNumber, <%=wn%> );
	onWeeknumber( oForm.yearlyWDay_weekNumber );
	selectItemOnValue( oForm.yearlyWDay_weekday, <%=wd%> );
	onSelectWeekday( oForm.yearlyWDay_weekday );
	<%

CASE "weekly"
	%>
	n = <%=CINT(aDateParse(0))%>
	selectItemOnValue( oForm.weeklyRepeat, n );
	<%
	DIM	a
	a = SPLIT(aDateParse(1),",")
	FOR EACH dd IN a
		SELECT CASE dd
		CASE "1"
			%>
			oForm.weeklySunday.checked = true;
			<%
		CASE "2"
			%>
			oForm.weeklyMonday.checked = true;
			<%
		CASE "3"
			%>
			oForm.weeklyTuesday.checked = true;
			<%
		CASE "4"
			%>
			oForm.weeklyWednesday.checked = true;
			<%
		CASE "5"
			%>
			oForm.weeklyThursday.checked = true;
			<%
		CASE "6"
			%>
			oForm.weeklyFriday.checked = true;
			<%
		CASE "7"
			%>
			oForm.weeklySaturday.checked = true;
			<%
		END SELECT
	NEXT
CASE "monthly"
	SELECT CASE sDateSubType
	CASE "dayn"
		%>
		n = <%=aDateParse(0)%>;
		selectItemOnValue( oForm.monthlyDayn_cycle, n );
		onSelectCycle( oForm.monthlyDayn_cycle );
		
		n = <%=aDateParse(1)%>;
		selectItemOnValue( oForm.monthlyDayn_index, n );
		onSelectMonthIndex( oForm.monthlyDayn_index );
		
		n = <%=aDateParse(2)%>;
		oForm.monthlyDayn_day.value = n;
		onDay( oForm.monthlyDayn_day );
		<%
		yy = YEAR(NOW)
		mm = MONTH(NOW) MOD CINT(aDateParse(0)) + 1
		dd = CINT(aDateParse(2))
		jd = jdateFromGregorian( dd, mm, yy )
		wd = weekdayFromJDate( jd )
		wn = (dd - 1) \ 7 + 1
		%>
		selectItemOnValue( oForm.yearlyWDay_weekNumber, <%=wn%> );
		onWeeknumber( oForm.yearlyWDay_weekNumber );
		selectItemOnValue( oForm.yearlyWDay_weekday, <%=wd%> );
		onSelectWeekday( oForm.yearlyWDay_weekday );
		<%

	case "wday"
		%>
		n = <%=aDateParse(0)%>;
		selectItemOnValue( oForm.monthlyWDay_cycle, n );
		onSelectCycle( oForm.monthlyWDay_cycle );
		
		n = <%=aDateParse(1)%>;
		selectItemOnValue( oForm.monthlyWDay_index, n );
		onSelectMonthIndex( oForm.monthlyWDay_index );
		
		n = <%=aDateParse(2)%>;
		selectItemOnValue( oForm.monthlyWDay_weekNumber, n );
		onWeeknumber( oForm.monthlyWDay_weekNumber );
		n = <%=aDateParse(3)%>;
		selectItemOnValue( oForm.monthlyWDay_weekday, n );
		onSelectWeekday( oForm.monthlyWDay_weekday );
		<%
		mm = CINT(aDateParse(1))
		wn = CINT(aDateParse(2))
		wd = CINT(aDateParse(3))
		jd = jdateFromWeeklyGregorian( wn, wd, mm, YEAR(NOW) )
		gregorianFromJDate dd, mm, yy, jd
		%>
		oForm.single_day.value = <%=dd%>;
		onDay( oForm.single_day );
		oForm.single_year.value = <%=yy%>;
		<%
	END SELECT
CASE "yearly"
	SELECT CASE sDateSubType
	CASE "dayn"
		%>
		n = <%=aDateParse(0)%>;
		selectItemOnValue( oForm.yearlyDayn_month, n );
		onSelectMonth( oForm.yearlyDayn_month );
		oForm.yearlyDayn_day.value = <%=aDateParse(1)%>;
		onDay( oForm.yearlyDayn_day );
		<%
		yy = YEAR(NOW)
		mm = CINT(aDateParse(0))
		dd = CINT(aDateParse(1))
		jd = jdateFromGregorian( dd, mm, yy )
		wd = weekdayFromJDate( jd )
		wn = (dd - 1) \ 7 + 1
		%>
		selectItemOnValue( oForm.yearlyWDay_weekNumber, <%=wn%> );
		onWeeknumber( oForm.yearlyWDay_weekNumber );
		selectItemOnValue( oForm.yearlyWDay_weekday, <%=wd%> );
		onSelectWeekday( oForm.yearlyWDay_weekday );
		<%

	case "wday"
		%>
		n = <%=aDateParse(0)%>;
		selectItemOnValue( oForm.yearlyWDay_month, n );
		onSelectMonth( oForm.yearlyWDay_month );
		n = <%=aDateParse(1)%>;
		selectItemOnValue( oForm.yearlyWDay_weekNumber, n );
		onWeeknumber( oForm.yearlyWDay_weekNumber );
		n = <%=aDateParse(2)%>;
		selectItemOnValue( oForm.yearlyWDay_weekday, n );
		onSelectWeekday( oForm.yearlyWDay_weekday );
		<%
		mm = CINT(aDateParse(0))
		wn = CINT(aDateParse(1))
		wd = CINT(aDateParse(2))
		jd = jdateFromWeeklyGregorian( wn, wd, mm, YEAR(NOW) )
		gregorianFromJDate dd, mm, yy, jd
		%>
		oForm.single_day.value = <%=dd%>;
		onDay( oForm.single_day );
		oForm.single_year.value = <%=yy%>;
		<%
	END SELECT
CASE "hebrew"
	SELECT CASE sDateSubType
	CASE "dayn"
		%>
		n = <%=aDateParse(0)%>;
		selectItemOnValue( oForm.hebrewDayn_month, n );
		onSelectHebrewMonth( oForm.hebrewDayn_month );
		oForm.hebrewDayn_day.value = <%=aDateParse(1)%>;
		onDay( oForm.hebrewDayn_day );
		<%
		yy = YEAR(NOW)
		mm = CINT(aDateParse(0))
		dd = CINT(aDateParse(1))
		jd = jdateFromGregorian( dd, mm, yy )
		wd = weekdayFromJDate( jd )
		wn = (dd - 1) \ 7 + 1
		%>
		selectItemOnValue( oForm.hebrewWDay_weekNumber, <%=wn%> );
		onWeeknumber( oForm.hebrewWDay_weekNumber );
		selectItemOnValue( oForm.hebrewWDay_weekday, <%=wd%> );
		onSelectWeekday( oForm.hebrewWDay_weekday );
		<%

	case "wday"
		%>
		n = <%=aDateParse(0)%>;
		selectItemOnValue( oForm.hebrewWDay_month, n );
		onSelectHebrewMonth( oForm.hebrewWDay_month );
		n = <%=aDateParse(1)%>;
		selectItemOnValue( oForm.hebrewWDay_weekNumber, n );
		onWeeknumber( oForm.hebrewWDay_weekNumber );
		n = <%=aDateParse(2)%>;
		selectItemOnValue( oForm.hebrewWDay_weekday, n );
		onSelectWeekday( oForm.hebrewWDay_weekday );
		<%
		mm = CINT(aDateParse(0))
		wn = CINT(aDateParse(1))
		wd = CINT(aDateParse(2))
		jd = jdateFromWeeklyGregorian( wn, wd, mm, YEAR(NOW) )
		gregorianFromJDate dd, mm, yy, jd
		%>
		oForm.single_day.value = <%=dd%>;
		onDay( oForm.single_day );
		oForm.single_year.value = <%=yy%>;
		<%
	END SELECT
CASE "keyword"
	%>
	selectItemOnValue( oForm.keyword_name, "<%=sDate%>" );
	<%
CASE "season"
	%>
	selectItemOnValue( oForm.season_name, "<%=sDate%>" );
	<%
END SELECT
%>
//-->
</script>

</body>

</html>
<%

SET g_oFSO = Nothing

%>