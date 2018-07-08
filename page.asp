<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

%>
<!-- library includes -->
<!--#include file="scripts\htmlformat.asp"-->
<!--#include file="scripts\htmlformatforum.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts/include_pagebody.asp"-->
<!--#include file="scripts\include_pagelist.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!-- actions -->
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_calendar.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\include_guid.asp"-->
<%



DIM g_sCommentKey
DIM g_sSessionGUID
g_sSessionGUID = makeGUID()

g_sCommentKey = g_sCookiePrefix & "_SendMessage"

Session(g_sCommentKey) = g_sSessionGUID






g_nTabID = Request("tab")


DIM g_nCategoryID
g_nCategoryID = Request("category")


g_nPageID = Request("page")


DIM	g_sPageList
g_sPageList = ""


DIM	g_sXPageTitle
g_sXPageTitle = ""


' the following are returned for the current page
DIM	g_sXFormat
DIM	g_sXTitle
DIM	g_dXDateModified
DIM	g_nXCategory
DIM	g_bXCategoryGroupBy
g_sXFormat = ""
g_sXTitle = ""

CONST kCategoryPrefix = "categories.Name"

IF 0 < INSTR(g_sTabSortDetails,kCategoryPrefix) THEN
	g_bXCategoryGroupBy = TRUE
ELSE
	g_bXCategoryGroupBy = FALSE
END IF



FUNCTION isNum( x )
	isNum = FALSE
	IF "" = CSTR(x) THEN EXIT FUNCTION
	isNum = ISNUMERIC(x)
END FUNCTION



SUB getPageInfo( nID )

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"Format, " _
		&		"Title, " _
		&		"Description, " _
		&		"Body, " _
		&		"DateModified, " _
		&		"Category, " _
		&		"Picture " _
		&	"FROM " _
		&		"pages " _
		&	"WHERE " _
		&		"pages.RID=" & nID & " " _
		&	";"
	
	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )

	IF NOT oRS IS Nothing THEN
	
		IF NOT oRS.EOF THEN
		
			DIM	oFormat
			DIM	oTitle
			DIM	oDescription
			DIM	oBody
			DIM	oDateModified
			DIM	oCategory
			DIM	oPicture
			
			SET oFormat = oRS.Fields("Format")
			SET oTitle = oRS.Fields("Title")
			SET oDescription = oRS.Fields("Description")
			SET oBody = oRS.Fields("Body")
			SET oDateModified = oRS.Fields("DateModified")
			SET oCategory = oRS.Fields("Category")
			SET oPicture = oRS.Fields("Picture")
			
			g_sXFormat = recString(oFormat)

			g_sXTitle = pagebody_formatString( g_sXFormat, recString(oTitle) )

			g_sXDescription = recString(oDescription)
			g_sXBody = recString(oBody)
			g_dXDateModified = recDate(oDateModified)
			g_nXCategory = recNumber(oCategory)
			g_sXPicture = recString(oPicture)
			
		
		END IF
	
		oRS.Close
		SET oRS = Nothing
	
	END IF



END SUB


DIM g_nPageList
g_nPageList = 0


SUB makePageList()

	DIM	dFuture
	dFuture = DATEADD("yyyy", 20, NOW)
	DIM	sDateFuture
	sDateFuture = MONTH(dFuture) & "/" & DAY(dFuture) & "/" & YEAR(dFuture)
	
	
	DIM oRS
	SET oRS = getPageListRS( g_nTabID )
	
	IF NOT oRS IS Nothing THEN
	
		g_nPageList = oRS.RecordCount
	
		IF 0 < oRS.RecordCount THEN
		
			DIM	oPageRID
			DIM	oFormat
			DIM	oTitle
			DIM	oCategory
			DIM	oCatName
			DIM	oCatShortName
			
			SET oPageRID = oRS.Fields("RID")
			SET oFormat = oRS.Fields("Format")
			SET oTitle = oRS.Fields("Title")
			SET oCategory = oRS.Fields("Category")
			SET oCatName = oRS.Fields("CatName")
			SET oCatShortName = oRS.Fields("CatShortName")
			
			DIM	sFormat
			DIM	sCatName
			DIM	sPrevCatName
			sPrevCatName = ""

			DIM	sTitle

			g_sPageList = g_sPageList & "<div class=""TabGroupVert"">" & vbCRLF
			g_sPageList = g_sPageList & "<ul>" & vbCRLF

			oRS.MoveFirst
			
			IF isNum(g_nPageID) THEN
				getPageInfo g_nPageID
			ELSE
				g_nPageID = recNumber(oPageRID)
				getPageInfo g_nPageID
			END IF
			g_nCategoryID = g_nXCategory
			g_sXPageTitle = g_sXTitle & " - "
			
			DO UNTIL oRS.EOF
				sFormat = recString(oFormat)
				IF g_bXCategoryGroupBy THEN
					sCatName = recString(oCatName)
					IF sPrevCatName <> sCatName THEN
						IF "" <> sPrevCatName THEN
							g_sPageList = g_sPageList & "<" & "/ul" & ">" & vbCRLF
						END IF
						sPrevCatName = sCatName
						g_sPageList = g_sPageList & "<li class=""GroupHeader"""
						IF isNum(g_nCategoryID) THEN
							IF CLNG(g_nCategoryID) = recNumber(oCategory) THEN
								g_sPageList = g_sPageList & " id=""GroupSelected"""
							END IF
						END IF
						g_sPageList = g_sPageList & "><span>"
						g_sPageList = g_sPageList & Server.HTMLEncode(sCatName)
						g_sPageList = g_sPageList & "</span></li>" & vbCRLF
						g_sPageList = g_sPageList & "<ul>" & vbCRLF
					END IF
				END IF
				sTitle = pagebody_formatString( sFormat, recString(oTitle) )
				g_sPageList = g_sPageList & "<li"
				IF isNum(g_nPageID) THEN
					IF CLNG(g_nPageID) = recNumber(oPageRID) THEN
						g_sPageList = g_sPageList & " class=""SelectedTab"""
					END IF
				END IF
				g_sPageList = g_sPageList & ">"
				g_sPageList = g_sPageList & "<a href=""page.asp"
				g_sPageList = g_sPageList & "?tab=" & g_nTabID
				g_sPageList = g_sPageList & "&category=" & recNumber(oCategory)
				g_sPageList = g_sPageList & "&page=" & recNumber(oPageRID)
				g_sPageList = g_sPageList & """><span>" & sTitle
				g_sPageList = g_sPageList & "</span></a></li>" & vbCRLF
				oRS.MoveNext
			LOOP

			IF g_bXCategoryGroupBy THEN
				g_sPageList = g_sPageList & "</ul>" & vbCRLF
			END IF
			g_sPageList = g_sPageList & "</ul>" & vbCRLF
			g_sPageList = g_sPageList & "</div>" & vbCRLF

		END IF
		
		oRS.Close
		SET oRS = Nothing
	
	END IF

END SUB


makePageList

IF "" <> g_sXPageTitle THEN
	g_sXPageTitle = g_sXPageTitle & g_sNavPageTitle & " - "
ELSE
	g_sXPageTitle = g_sNavPageTitle & " - "
END IF





%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & g_sXPageTitle & Server.HTMLEncode(g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Page</title>
<%
END IF
%>
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" type="text/css" href="<%=g_sTheme%>style.css">
<%
IF LCASE(g_sTabName) = LCASE(g_sSiteTabNewsletters) THEN
%>
<link rel="stylesheet" type="text/css" href="newsletter.css">
<link rel="stylesheet" href="<%=remindCSS()%>" type="text/css">
<%
END IF
%>
<style type="text/css">





.DateModified
{
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: x-small;
	color: silver;
	margin-top: 3em;
}



.EmailNotification
{
	font-family: sans-serif;
	font-size: small;
	text-align: right;
}






</style>
<script type="text/javascript">

var g_sSessionGUID = "<%=g_sSessionGUID%>";

</script>
</head>

<body>
<!--#include file="scripts\page_begin.asp"-->
<div class="noprint">
<!--#include file="config\header.asp"-->
</div>
<%
IF 1 < g_nPageList THEN
Response.Write g_sPageList
END IF


%>

<table border="0" cellpadding="4" id="PageContainer">
<tr>
<td>

<%
pagebody_saveFuncs

outputFormattedPageInfo g_sXFormat, g_sXDescription, g_sXBody, g_sXPicture

pagebody_restoreFuncs


%>

</td>
</tr>
</table>
<%


IF g_dLocalPageDateModified < g_dXDateModified THEN g_dLocalPageDateModified = g_dXDateModified


%>
<div class="DateModified">Updated: <%=g_dXDateModified%></div>
<%
IF LCASE(g_sTabName) = LCASE(g_sSiteTabNewsletters) THEN
%>
<div class="EmailNotification"><a href="notification_newsletter.asp">Sign up for email notification</a></div>
<%
END IF
%>
<!--#include file="scripts\page_end.asp"-->


</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
