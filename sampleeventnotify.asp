<%@ Language=VBScript 
%>
<%
OPTION EXPLICIT


IF FALSE THEN
%>

<head>
<title>sample event notify</title>
<base target="_top">
</head>

<%
END IF





DIM	sMon
DIM	m
sMon = Request("m")
IF 0 < LEN(sMon) THEN
	m = CINT(sMon)
ELSE
	m = 3
END IF




DIM	g_sTabSortDetails
IF "mssql" = dbVendor() THEN
g_sTabSortDetails = "(ISNULL(categories.SortName,'')+categories.Name), " _
	&	"(ISNULL( pages.SortName,'')+pages.PageName), " _
	&	"pages.Title"
ELSEIF "mysql" = dbVendor() THEN
g_sTabSortDetails = "CONCAT(IFNULL(categories.SortName,''),categories.Name), " _
	&	"CONCAT(IFNULL( pages.SortName,''),pages.PageName), " _
	&	"pages.Title"
END IF




%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\include_forum.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts/rss.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_motorcycles.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!--#include file="scripts\include_announce.asp"-->
<!--#include file="scripts/include_pagelist.asp"-->
<!--#include file="scripts/include_pagebody.asp"-->
<!--#include file="scripts\include_emailnotificationformat.asp"-->
<%




DIM	dToday
dToday = findNearestNotificationDate()

DIM	oSMTP
oSMTP = ""


g_htmlFormat_pictureFunc = "remindPicture"
g_sProcessAnnouncementsPictureFunc = "f_processNotifyAnnouncementsPicture"

DIM	sBody
sBody = notify_buildHtmlBody( oSMTP, dToday )
Response.Write sBody





%>
<!--#include file="scripts\include_db_end.asp"-->
<%
%>