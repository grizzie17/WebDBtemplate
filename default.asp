<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT



%>
<!-- library includes -->
<!--#include file="scripts\dateutils.asp"-->
<!--#include file="scripts\htmlformat.asp"-->
<!--#include file="scripts\include_announce.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_emailnotificationformat.asp"-->
<!--#include file="scripts\include_forum.asp"-->
<!--#include file="scripts\include_pagebody.asp"-->
<!--#include file="scripts\include_pagelist.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<!--#include file="scripts\include_rssannounce.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts\remind_files.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<!--#include file="scripts\rss.asp"-->
<!--#include file="scripts\rssweather.asp"-->
<!-- actions -->
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_mobile.asp"-->
<%



DIM g_dLastVisit
DIM	sLastVisit
g_dLastVisit = CDATE("1/1/2000")

sLastVisit = Session( g_sCookiePrefix & "_LastVisit" )
IF "" <> CSTR(sLastVisit) THEN
	g_dLastVisit = CDATE(sLastVisit)
ELSE
	sLastVisit = Request.Cookies( g_sCookiePrefix & "_LastVisit" )
	IF "" <> CSTR(sLastVisit) THEN
		g_dLastVisit = DATEADD("h", -4, CDATE(sLastVisit))
	ELSE
		g_dLastVisit = DATEADD("d", -6, NOW)
	END IF
END IF


Session( g_sCookiePrefix & "_LastVisit" ) = CSTR(g_dLastVisit)

Response.Cookies( g_sCookiePrefix & "_LastVisit" ) = NOW
Response.Cookies( g_sCookiePrefix & "_LastVisit" ).Expires = DateAdd( "d", 90, NOW )









SUB outputDynamic( sName, sURL, sCacheFile, nTimer, sInterval, nIntValue, sBreakInterval )

	DIM	sPath
	DIM	oFile
	DIM	sContent
	sPath = cache_checkFile( "site", sCacheFile, sInterval, nIntValue, sBreakInterval )
	IF 0 < LEN(sPath) THEN
		SET oFile = g_oFSO.OpenTextFile( sPath, 1 )
		IF NOT Nothing IS oFile THEN
			sContent = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			Response.Write "<div id=""" & sName & "content"" class=""dynamic"">" & vbCRLF
			Response.Write sContent
			Response.Write "</div>" & vbCRLF
		END IF
	ELSE
		Response.Write "" _
			&	"<script type=""text/javascript"">" & vbCRLF _
			&	"function dynamicload" & sName & "() {" & vbCRLF _
			&		"ajaxPage('" & sURL & "', '" & sName & "content');" & vbCRLF _
			&	"}" & vbCRLF _
			&	"function timerajax" & sName & "() {" & vbCRLF _
			&		"setTimeout('dynamicload" & sName & "()', " & nTimer & " * 1000);" & vbCRLF _
			&	"}" & vbCRLF _
			&	"</script>" & vbCRLF
		Response.Write "<div id=""" & sName & "content"" class=""dynamic"">" & vbCRLF
		sPath = cache_filepath( "site", sCacheFile )
		IF g_oFSO.FileExists( sPath ) THEN
			SET oFile = g_oFSO.OpenTextFile( sPath, 1 )
			IF NOT Nothing IS oFile THEN
				sContent = oFile.ReadAll
				oFile.Close
				SET oFile = Nothing
				Response.Write sContent
			END IF
		ELSE
			Response.Write "<img src=""images/loading.gif"" alt=""Loading ..."">"
		END IF
		Response.Write "" _
			&	"<script type=""text/javascript"">" & vbCRLF _
			&		"if (window.addEventListener)" & vbCRLF _
			&			"window.addEventListener(""load"", timerajax" & sName & ", false);" & vbCRLF _
			&		"else if (window.attachEvent)" & vbCRLF _
			&			"window.attachEvent(""onload"", timerajax" & sName & ");" & vbCRLF _
			&		"else" & vbCRLF _
			&			"setTimeout(""timerajax" & sName & "()"", 1.0 * 1000);" & vbCRLF _
			&		"</script>" & vbCRLF
		Response.Write "</div>" & vbCRLF
	END IF

END SUB






SUB loadRemindLastModified
	DIM	oFile
	SET oFile = cache_openTextFile( "remind", "RemindLastModified.txt", "m", 1, "y" )
	IF NOT oFile IS Nothing THEN
		DIM	sDate
		sDate = oFile.ReadLine
		oFile.Close
		SET oFile = Nothing
		IF "" <> sDate THEN
			dRemindLastModified = CDATE(sDate)
		END IF
	END IF
END SUB
loadRemindLastModified

DIM	sKeywords
	sKeywords = Server.HTMLEncode(g_sSiteName&","&g_sSiteCity&","&g_sShortSiteName&","&g_sMetaKeywords)


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & Server.HTMLEncode(g_sSiteName & " - " & g_sShortSiteName & " - " & g_sSiteCity ) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Home - Default</title>
<%
END IF
%>
<meta name="keywords" content="<%=sKeywords%>">
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<script type="text/javascript" language="javascript" src="scripts/ajax.js"></script>
<link rel="stylesheet" type="text/css" href="default.css">
<link rel="stylesheet" type="text/css" href="<%=g_sTheme%>style.css">
<link rel="stylesheet" type="text/css" href="scripts/rssforum.css">
<link rel="stylesheet" type="text/css" href="<%=remindCSS()%>">
<script type="text/javascript">
<!--
var sThisURL = unescape(window.location.pathname);

function doFramesBuster()
{
	if ( top != self )
	{
		if ( top.location.replace )
			top.location.replace( sThisURL );
		else
			setTimeout( "top.location.href = sThisURL", 1.5*1000 );
	}
}

if ( "MSIE" == navigator.appName  ||  -1 < (navigator.appName).indexOf("Microsoft") )
	doFramesBuster();



//-->
</script>
<%
makeWeatherStyles
%>
</head>

<body onload="doFramesBuster()" class="homepage">

<!--#include file="scripts\page_begin.asp"-->
<!--#include file="config\header.asp"-->
<input type="hidden" id="refreshed" value="no">
<script type="text/javascript">
onload=function()
{
	var e=document.getElementById("refreshed");
	if (e.value=="no")
	{
		e.value="yes";
	}
	else
	{
		e.value="no";
		location.reload();
	}
}
</script>
<%
IF NOT g_bMobileBrowser THEN
%>
<!--#include file="app_data\layout\default_body.asp"-->
<%
ELSE
%>
<!--#include file="app_data\layout\default_body_mobile.asp"-->
<%
END IF
%>

<hr>
<%
IF g_dLocalPageDateModified < dRemindLastModified THEN g_dLocalPageDateModified = dRemindLastModified
%>
<!--#include file="scripts\page_end.asp"-->
<p style="font-size: x-small; color: silver">Last Visit: <%=DATEADD("h", g_nServerTimeZoneOffset, g_dLastVisit)%></p>
<%

IF 0 < LEN(g_sUpcomingEventsSchedule) THEN
	outputDynamic "emailnotice", "loadnotifications.asp", "emailnotice.htm", "1.0", "h", 24, "d"
END IF

%>

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
