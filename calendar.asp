<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT





DIM	sMon
DIM	m
sMon = Request("m")
IF 0 < LEN(sMon) THEN
	m = CINT(sMon)
ELSE
	m = 3
END IF

DIM sPreface
sPreface = Request("p")
IF "" = sPreface THEN
	sPreface = "group"
END IF


DIM g_sNavigateTabLabel
g_sNavigateTabLabel = "(tab|calendar)"


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts/include_db_begin.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/remind_liststyles.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!--#include file="scripts/include_navtabs.asp"-->
<!--#include file="scripts/include_pagelist.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\tab_tools.asp"-->
<!--#include file="scripts/include_pagebody.asp"-->
<%



'DIM g_nPageID
g_nPageID = Request("page")


DIM	g_sPageList
g_sPageList = ""


DIM	g_sXPageTitle
g_sXPageTitle = ""

' the following are returned for the current page
DIM	g_sXFormat
DIM	g_sXTitle
'DIM	g_sXDescription
'DIM	g_sXBody
DIM	g_dXDateModified
DIM	g_nXCategory
DIM	g_bXCategoryGroupBy
'DIM	g_sXPicture
g_sXFormat = ""
g_sXTitle = ""
g_sXDescription = ""
g_sXBody = ""
g_sXPicture = ""

CONST kCategoryPrefix = "categories.Name"

IF 0 < INSTR(g_sTabSortDetails,kCategoryPrefix) THEN
	g_bXCategoryGroupBy = TRUE
ELSE
	g_bXCategoryGroupBy = FALSE
END IF





%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & g_sNavPageTitle & " - " & Server.HTMLEncode(g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Calendar</title>
<%
END IF
%>
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<link rel="stylesheet" href="<%=remindCSS()%>" type="text/css">
<style type="text/css">

.TabGroupVert
{
	width: 8em;
	float: left;
}


.hiddenlist
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: x-small;
}

.hiddenlist a:link, a:active, a:visited
{
	color: #6699FF;
}


.remindcsslist
{
	margin-left: 1em;
	padding-left: 0;
}

td ul.htmlformat
{
	margin-left: 1em;
	padding-left: 0;
}






</style>
</head>

<body>

<!--#include file="scripts\page_begin.asp"-->
<div class="noprint">
<!--#include file="config\header.asp"-->
</div>
<%


CLASS CTabFormatGray

	PROPERTY GET colorBackground()
		colorBackground = "#999999"
	END PROPERTY
	
	PROPERTY GET colorTab()
		colorTab = "#CCCCCC"
	END PROPERTY
	
	PROPERTY GET colorTabSelected()
		colorTabSelected = "#FFFFFF"
	END PROPERTY
	
	PROPERTY GET classTabGroup()
		classTabGroup = "SubNavigation"
	END PROPERTY
	
	PROPERTY GET classTabGroupInverted()
		classTabGroupInverted = "SubNavigationInverted"
	END PROPERTY
	
	PROPERTY GET classTabGroupVert()
		classTabGroupVert = "TabGroupVert"
	END PROPERTY
	
	PROPERTY GET classTab()
		classTab = "shoppingtab"
	END PROPERTY
	
	PROPERTY GET classSelected()
		classSelected = "SelectedTab"
	END PROPERTY
	
	PROPERTY GET alignTabHoriz()
		alignTabHoriz = "center"
	END PROPERTY
	
	PROPERTY GET imageTL()
		imageTL = "images/pie_tl_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageTR()
		imageTR = "images/pie_tr_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageBL()
		imageBL = "images/pie_bl_gray.gif"
	END PROPERTY

	PROPERTY GET imageBR()
		imageBR = "images/pie_br_gray.gif"
	END PROPERTY
	
END CLASS


CLASS CTabMemberData

	PRIVATE m_aData
	PRIVATE m_aSplit
	PRIVATE m_i
	PRIVATE m_sURL

	PRIVATE SUB Class_Initialize
		m_sURL = ""
		m_i = 0
	END SUB

	
	PROPERTY GET RecordCount()
		RecordCount = UBOUND(m_aData) + 1
	END PROPERTY
	
	PROPERTY GET EOF()
		IF m_i <= UBOUND(m_aData) THEN
			EOF = FALSE
		ELSE
			EOF = TRUE
		END IF
	END PROPERTY
	
	SUB MoveFirst()
		m_i = 0
		privParse
	END SUB
	
	SUB MoveNext()
		m_i = m_i + 1
		privParse
	END SUB
	
	FUNCTION IsCurrent( x )
		IF m_i = x THEN
			IsCurrent = TRUE
		ELSE
			IsCurrent = FALSE
		END IF
	END FUNCTION
	
	PROPERTY GET HREF()
		HREF = m_sURL & m_aSplit(1)
	END PROPERTY
	
	PROPERTY GET TabLabel()
		TabLabel = m_aSplit(0)
	END PROPERTY
	
	'----------------
	
	PRIVATE SUB privParse()
		IF NOT EOF() THEN
			m_aSplit = SPLIT( m_aData(m_i), vbTAB, -1, vbTextCompare )
		END IF
	END SUB
	
	PROPERTY LET Data( a )
		m_aData = a
	END PROPERTY
	
	PROPERTY LET URL( s )
		m_sURL = s
	END PROPERTY
	
END CLASS


	DIM	oTabGen
	DIM	oTabData
	DIM	oTabFormat
	DIM	g_sPage
	
	g_sPage = Request("cat")
	
	SET oTabGen = NEW CTabGenerate
	SET oTabData = NEW CTabMemberData
	SET oTabFormat = NEW CTabFormatGray
	
	
	DIM	aCatList
	aCatList = SPLIT( g_sCalendarList, vbLF )

	DIM	aHiddenList
	DIM	sHiddenItem
	DIM	aHiddenStuff
	aHiddenList = SPLIT( g_sCalendarHiddenList, vbLF )
	FOR EACH sHiddenItem IN aHiddenList
		aHiddenStuff = SPLIT( sHiddenItem, vbTAB )
		IF aHiddenStuff(1) = g_sPage THEN
			REDIM PRESERVE aCatList(UBOUND(aCatList)+1)
			aCatList(UBOUND(aCatList)) = sHiddenItem
			EXIT FOR
		END IF
	NEXT 'sHiddenItem

	DIM	g_sFile
	DIM	g_sPageTitle
	DIM	g_nIndex
	DIM	nLen
	DIM	aFileSplit

	g_sFile = ""
	g_sPageTitle = ""
	g_nIndex = -1

	FOR nLen = 0 TO UBOUND(aCatList)
		aFileSplit = SPLIT( aCatList(nLen), vbTAB, -1, vbTextCompare )
		IF LCASE(g_sPage) = LCASE(aFileSplit(1)) THEN
			g_sPageTitle = aFileSplit(2)
			g_nIndex = nLen
			EXIT FOR
		END IF
	NEXT 'nLen
	IF g_nIndex < 0 THEN
		g_nIndex = UBOUND(aCatList)+1
		REDIM PRESERVE aCatList(g_nIndex)
		aCatList(g_nIndex) = g_sPage & vbTAB & LCASE(g_sPage) & vbTAB & g_sPage & " Events"
	END IF
	IF "" = g_sPageTitle THEN
		aFileSplit = SPLIT( aCatList(g_nIndex), vbTAB, -1, vbTextCompare )
		g_sPageTitle = aFileSplit(2)
	END IF
	IF "All" = g_sPageTitle THEN
		g_sPageTitle = "Events / Rides"
	END IF

	
	oTabData.Data = aCatList
	oTabData.URL = "calendar.asp?m=" & sMon & "&p=" & sPreface & "&cat="
	
	SET oTabGen.TabData = oTabData
	SET oTabGen.TabFormat = oTabFormat
	oTabGen.MaxCols = 10
	oTabGen.Current = g_nIndex
	oTabGen.makeTabs
	
	
	
	SET oTabFormat = Nothing
	SET oTabData = Nothing
	SET oTabGen = Nothing


%>
<table class="noprint" border="1" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" style="border-collapse: collapse" width="200" id="table1" bordercolor="#FFCC00" align="right">
  <tr>
    <th bgcolor="#FFDD99" nowrap align="center">Display Months</th>
  </tr>
  <tr>
    <th nowrap align="center">
      <a href="calendar.asp?m=3&cat=<%=g_sPage%>&p=<%=sPreface%>">3</a>, 
      <a href="calendar.asp?m=6&cat=<%=g_sPage%>&p=<%=sPreface%>">6</a>, 
      <a href="calendar.asp?m=9&cat=<%=g_sPage%>&p=<%=sPreface%>">9</a>,<br>
      <a href="calendar.asp?m=12&cat=<%=g_sPage%>&p=<%=sPreface%>">12</a>, 
      <a href="calendar.asp?m=18&cat=<%=g_sPage%>&p=<%=sPreface%>">18</a>, 
      <a href="calendar.asp?m=24&cat=<%=g_sPage%>&p=<%=sPreface%>">24</a>,<br>
      <a href="calendar.asp?m=36&cat=<%=g_sPage%>&p=<%=sPreface%>">36</a>, 
      <a href="calendar.asp?m=48&cat=<%=g_sPage%>&p=<%=sPreface%>">48</a>, 
      <a href="calendar.asp?m=60&cat=<%=g_sPage%>&p=<%=sPreface%>">60</a></th>
  </tr>
  <tr>
    <th bgcolor="#FFDD99" nowrap align="center">Printer Friendly</th>
  </tr>
<%
IF g_bCalendarPrefaceOption THEN
%>
  <tr>
    <td bgcolor="#FFDD99" nowrap align="center">
    <font size="1">
    <%
    
    Response.Write "<a href=""calendar.asp?m=" & m & "&cat=" & g_sPage & "&p="
    IF "group" = sPreface THEN
    	Response.Write "none"">Group Ride Preface"
    ELSE
    	Response.Write "group"">No Group Ride Preface"
    END IF
    Response.Write "</a>"
    
    %></font>
    </td>
  </tr>
<%
END IF
%>
  <tr>
    <th nowrap align="center">
      <a target="_blank" href="calendar_print.asp?m=1&cat=<%=g_sPage%>&p=<%=sPreface%>">1</a>, 
      <a target="_blank" href="calendar_print.asp?m=2&cat=<%=g_sPage%>&p=<%=sPreface%>">2</a>, 
      <a target="_blank" href="calendar_print.asp?m=3&cat=<%=g_sPage%>&p=<%=sPreface%>">3</a>, 
      <a target="_blank" href="calendar_print.asp?m=4&cat=<%=g_sPage%>&p=<%=sPreface%>">4</a>, 
      <a target="_blank" href="calendar_print.asp?m=6&cat=<%=g_sPage%>&p=<%=sPreface%>">6</a>, 
      <a target="_blank" href="calendar_print.asp?m=9&cat=<%=g_sPage%>&p=<%=sPreface%>">9</a>,<br>
      <a target="_blank" href="calendar_print.asp?m=12&cat=<%=g_sPage%>&p=<%=sPreface%>">12</a>, 
      <a target="_blank" href="calendar_print.asp?m=18&cat=<%=g_sPage%>&p=<%=sPreface%>">18</a>, 
      <a target="_blank" href="calendar_print.asp?m=24&cat=<%=g_sPage%>&p=<%=sPreface%>">24</a>,<br>
      <a target="_blank" href="calendar_print.asp?m=36&cat=<%=g_sPage%>&p=<%=sPreface%>">36</a>, 
      <a target="_blank" href="calendar_print.asp?m=48&cat=<%=g_sPage%>&p=<%=sPreface%>">48</a>, 
      <a target="_blank" href="calendar_print.asp?m=60&cat=<%=g_sPage%>&p=<%=sPreface%>">60</a><br>
		<font size="1">Opens a New window</font></th>
  </tr>
  <tr>
    <th bgcolor="#FFDD99" nowrap align="center">Month at a Glance</th>
  </tr>
  <tr>
    <th nowrap align="center">
      <a target="_blank" href="calendar_mag.asp?m=1&cat=<%=g_sPage%>&p=<%=sPreface%>">1</a>, 
      <a target="_blank" href="calendar_mag.asp?m=2&cat=<%=g_sPage%>&p=<%=sPreface%>">2</a>, 
      <a target="_blank" href="calendar_mag.asp?m=3&cat=<%=g_sPage%>&p=<%=sPreface%>">3</a>, 
      <a target="_blank" href="calendar_mag.asp?m=4&cat=<%=g_sPage%>&p=<%=sPreface%>">4</a>, 
      <a target="_blank" href="calendar_mag.asp?m=6&cat=<%=g_sPage%>&p=<%=sPreface%>">6</a>, 
      <a target="_blank" href="calendar_mag.asp?m=9&cat=<%=g_sPage%>&p=<%=sPreface%>">9</a>,<br>
      <a target="_blank" href="calendar_mag.asp?m=12&cat=<%=g_sPage%>&p=<%=sPreface%>">12</a>, 
      <a target="_blank" href="calendar_mag.asp?m=18&cat=<%=g_sPage%>&p=<%=sPreface%>">18</a>, 
      <a target="_blank" href="calendar_mag.asp?m=24&cat=<%=g_sPage%>&p=<%=sPreface%>">24</a>,<br>
      <a target="_blank" href="calendar_mag.asp?m=36&cat=<%=g_sPage%>&p=<%=sPreface%>">36</a>, 
      <a target="_blank" href="calendar_mag.asp?m=48&cat=<%=g_sPage%>&p=<%=sPreface%>">48</a>, 
      <a target="_blank" href="calendar_mag.asp?m=60&cat=<%=g_sPage%>&p=<%=sPreface%>">60</a><br>
		<font size="1">Opens a New window</font></th>
  </tr>
	<tr>
		<th bgcolor="#FFDD99" align="center">
		<a name="register"></a>Register for Email</th>
	</tr>
	<tr>
		<td>
		<p>Register for email notification of <a href="notification_upcoming.asp">upcoming calendar events</a>.</td>
	</tr>
<%


pagebody_saveFuncs
outputMultiplePages FALSE	' FALSE to suppress generation of the table header
pagebody_restoreFuncs




%>
	<tr>
		<th bgcolor="#FFDD99" align="center">Calendar Key</th>
	</tr>
	<tr>
		<td>
		We use different colors, text-sizes, and fonts to denote different types 
		of events. In general, if something in the calendar is underlined then 
		it is a link.
<%


remind_liststyles



%>
		</td>
	</tr>
	<tr>
	<td class="hiddenlist">
	<%
	FOR EACH sHiddenItem IN aHiddenList
		aHiddenStuff = SPLIT( sHiddenItem, vbTAB )
		Response.Write "<a href=""calendar.asp?m=" & sMon & "&p=" & sPreface & "&cat=" & aHiddenStuff(1) & """>"
		Response.Write Server.HTMLEncode(aHiddenStuff(0))
		Response.Write "</a><br>" & vbCRLF
	NEXT 'sHiddenItem
	%>
	</td>
	</tr>
</table>
<%
IF Response.Buffer THEN Response.Flush
%>

<center>
<table border="1" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" style="border-collapse: collapse" id="table1" bordercolor="#FFCC00">
	<tr>
		<th bgcolor="#FFDD99" align="center"><%=Server.HTMLEncode(g_sPageTitle)%> 
		(<%=m%> Months)</th>
	</tr>
	<tr>
		<td align="left">
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



nDateBegin = jdateFromVBDate( NOW )
'nDateEnd = nDateBegin + 28 * 3 - 1
'nDateEnd = -14	'pending

gHtmlOption_encodeEmailAddresses = TRUE

DIM	sSavePictureFunc
sSavePictureFunc = g_htmlFormat_pictureFunc
g_htmlFormat_pictureFunc = "remindPicture"


SET oCalendar = loadRemindFiles2( nDateBegin, nDateEnd, _
							g_sPage, "holiday,usa,none", TRUE, dNow, "list" )
IF NOT Nothing IS oCalendar THEN

	' change all events that are past
	
	DIM	oEvents
	DIM	oXML
	SET oXML = oCalendar.xmldom
	
	SET oEvents = oXML.documentElement.selectNodes("/calendar/event[date < '" & logDateFromJDate( jdateFromVBDate( dNow ) ) & "']")
	IF NOT( Nothing IS oEvents) THEN
	
		'Response.Write "<p>Got some events</p>"
		DIM	oEvent
		DIM	oStyle
		DIM	oText
		FOR EACH oEvent IN oEvents
			'Response.Write "<p>oEvent</p>"
			SET oStyle = oEvent.selectSingleNode("style")
			IF NOT Nothing IS oStyle THEN
				oStyle.text = "RmdPast"
			ELSE
				SET oText = oCalendar.xmldom.createTextNode( "RmdPast" )
				SET oStyle = oCalendar.xmldom.createNode( 1, "style", "" )
				oStyle.appendChild( oText )
				oEvent.appendChild( oStyle )
			END IF
		NEXT 'oEvent
	END IF

	'Load the XSL
	DIM	oXSL
	SET oXSL = remindLoadXmlFile( Server.MapPath("scripts/remind.xslt") )
	'SET oXSL = Server.CreateObject("msxml2.DOMDocument")
	'oXSL.async = false
	'oXSL.load(Server.MapPath("scripts/remind.xslt"))
	
	Response.Write oXML.transformNode(oXSL)

END IF


%> 
</td>
	</tr>
</table>
<p style="color:#999999; font-family: sans-serif; font-size: xx-small;">Calendar 
Updated: <%=DATEADD("h", g_nServerTimeZoneOffset, dRemindLastModified)%></p>

</center>

<%

IF g_dLocalPageDateModified < dRemindLastModified THEN g_dLocalPageDateModified = dRemindLastModified

%>


<!--#include file="scripts\page_end.asp"-->

</body>

</html>
<!--#include file="scripts/include_db_end.asp"-->
