<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

	'"Motorcyclist.xml" & vbTAB & "http://www.motorcyclistonline.com/rss/nonstandard_rss.xml", _

DIM	nNewsIndex
DIM	aNews
aNews = ARRAY( _
	"Motorcycle_Cruiser-All.xml" & vbTAB & "http://www.motorcyclecruiser.com/rss/all/index.html" & vbTAB & "www.MotorcycleCruiser.com", _
	"Motorcycle_Cruiser-Street_Survival.xml" & vbTAB & "http://www.motorcyclecruiser.com/rss/articles/streetsurvival/index.html" & vbTAB & "www.motorcyclecruiser.com", _
	"Motorcyclist.xml" & vbTAB & "http://www.motorcyclistonline.com/rss/articles/newsandupdates" & vbTAB & "www.MotorcyclistOnline.com", _
	"Motorcyclist_Features.xml" & vbTAB & "http://www.motorcyclistonline.com/rss/articles/features" & vbTAB & "www.MotorcyclistOnline.com", _
	"MC_Safety_Tips.xml" & vbTAB & "http://www.msgroup.org/forums/mtt/rss.asp?MEMBER_ID=45&ChkID=4e4bc071e3cf", _
	"-", _
	"Bible_Gateway.xml" & vbTAB & "http://www.biblegateway.com/usage/votd/rss/votd.rdf?31", _
	"Got_Questions.xml" & vbTAB & "http://www.gotquestions.org/questweek.xml&t1=" & vbTAB & "www.GotQuestions.org", _
	"TurningPoint.xml" & vbTAB & "http://media.salemwebnetwork.com/rss/oneplace/ministries/broadcastarchives/29.xml", _
	"The_Baptist_Standard.xml" & vbTAB & "http://feeds.feedburner.com/TheBaptistStandard", _
	"Answers_In_Genesis-Daily.xml" & vbTAB & "http://feeds.feedburner.com/AIGDaily" & vbTAB & "www.AnswersInGenesis.org", _
	"Answers_Research_Journal.xml" & vbTAB & "http://feeds.feedburner.com/ARJ" & vbTAB & "www.AnswersInGenesis.org/ARJ", _
	"Devotions.xml" & vbTAB & "http://www.onenewsnow.com/rss/devotions.aspx" & vbTAB & "www.OneNewsNow.com", _
	"-", _
	"One_News_Now.xml" & vbTAB & "http://www.onenewsnow.com/rss/rss_allsections.aspx" & vbTAB & "www.OneNewsNow.com", _
	"CSM-Top_Stories.xml" & vbTAB & "http://www.csmonitor.com/rss/top.rss" & vbTAB & "www.csmonitor.com", _
	"WND-Page1.xml" & vbTAB & "http://wnd.com/?ol=0&fa=PAGE.rss" & vbTAB & "www.WND.com", _
	"WND-Breaking_News.xml" & vbTAB & "http://www.wnd.com/index.php?fa=PAGE.rss&childFilter_10=1" & vbTAB & "www.WND.com", _
	"WND-Commentary.xml" & vbTAB & "http://www.worldnetdaily.com/index.php?fa=PAGE.rss&title=Commentary" & vbTAB & "www.WND.com", _
	"-", _
	"Dave_Ramsey.xml" & vbTAB & "http://www.daveramsey.com/rss/?fuseaction=dspGetFeed&strFeeds=MyTmmo,DaveSays,StupidTax,PressRelease" & vbTab & "www.DaveRamsey.com", _
	"-", _
	"Msdn.xml" & vbTAB & "http://www.microsoft.com/feeds/msdn/en-us/rss.xml" & vbTAB & "msdn.microsoft.com", _
	"Expression_Forum.xml" & vbTAB & "http://social.expression.microsoft.com/Forums/en-US/web/threads?outputAs=atom" & vbTAB & "social.expression.microsoft.com/Forums/en-US/web/threads", _
	"Expression_News.xml" & vbTAB & "http://www.microsoft.com/feeds/Expression/en-us/Expression_en_us.xml" & vbTAB & "expression.microsoft.com", _
	"Expression_Whats_Fresh.xml" & vbTAB & "http://www.microsoft.com/feeds/Expression/en-us/Expression_Fresh.xml" & vbTAB & "expression.microsoft.com", _
	"Expression_Features.xml" & vbTAB & "http://www.microsoft.com/feeds/Expression/en-us/Expression_Features.xml" & vbTAB & "expression.microsoft.com", _
	"Expression_Tips.xml" & vbTAB & "http://services.social.microsoft.com/feeds/feed/ExpressionProductTips" & vbTAB & "expression.microsoft.com", _
	"Office_Hours_Columns.xml" & vbTAB & "http://office.microsoft.com/download/afile.aspx?assetid=HX102367541033" & vbTAB & "office.microsoft.com/en-us/help/CH102264241033.aspx", _
	"PC_World.xml" & vbTAB & "http://feeds.pcworld.com/pcworld/latestnews", _
	"Windows_IT_Pro.xml" & vbTAB & "http://feeds.penton.com/windowsitpro/wininfo", _
	"-", _
	"Cnet_News-20.xml" & vbTAB & "http://news.com.com/2547-1_3-0-20.xml", _
	"Fox_Technology.xml" & vbTAB & "http://feeds.foxnews.com/foxnews/scitech", _
	"Yahoo_Science.xml" & vbTAB & "http://rss.news.yahoo.com/rss/science", _
	"Yahoo_Tech.xml" & vbTAB & "http://rss.news.yahoo.com/rss/tech", _
	"Vista_News.xml" & vbTAB & "http://www.vistanews.com/rss/index.xml", _
	"WXP_News.xml" & vbTAB & "http://www.wxpnews.com/rss/index.xml", _
	"Win7News.xml" & vbTAB & "http://www.win7news.net/rss/index.xml" & vbTAB & "www.win7news.net", _
	"Zdnet_Web_News.xml" & vbTAB & "http://news.zdnet.com/2509-9588_22-0-10.xml", _
	"Zdnet_News.xml" & vbTAB & "http://zdnet.com.com/2509-1_22-0-5.xml", _
	"CNN_Science.xml" & vbTAB & "http://rss.cnn.com/rss/cnn_space.rss", _
	"CNN_Tech.xml" & vbTAB & "http://rss.cnn.com/rss/cnn_tech.rss", _
	"How_Stuff_Works.xml" & vbTAB & "http://feeds.howstuffworks.com/DailyStuff" & vbTAB & "www.HowStuffWorks.com", _
	"How_Stuff_Works-Computers.xml" & vbTAB & "http://feeds.feedburner.com/HowstuffworksComputerstuffDailyRssFeed" & vbTAB & "www.HowStuffWorks.com", _
	"How_Stuff_Works-Future_Cars.xml" & vbTAB & "http://feeds.feedburner.com/HowstuffworksFutureCarsFromConsumerGuideAuto" & vbTAB & "www.HowStuffWorks.com", _
	"How_Stuff_Works-Science.xml" & vbTAB & "http://feeds.feedburner.com/hsw-science" & vbTAB & "www.HowStuffWorks.com", _
	"-", _
	"NetFlix_Queue.xml" & vbTAB & "http://rss.netflix.com/QueueRSS?id=P1414328699671123226356931042167502A" & vbTAB & "www.NetFlix.com", _
	"NetFlix_Queue-Now.xml" & vbTAB & "http://rss.netflix.com/QueueEDRSS?id=P1414328699671123226356931042167502A" & vbTAB & "www.NetFlix.com", _
	"NetFlix_Movies-At-Home.xml" & vbTAB & "http://rss.netflix.com/AtHomeRSS?id=P1414328699671123226356931042167502A" & vbTAB & "www.NetFlix.com", _
	"NetFlix_Recommendations.xml" & vbTAB & "http://rss.netflix.com/RecommendationsRSS?id=P1414328699671123226356931042167502A" & vbTAB & "www.NetFlix.com", _
	"NetFlix_New-Now.xml" & vbTAB & "http://www.netflix.com/NewWatchInstantlyRSS" & vbTAB & "www.NetFlix.com", _
	"-", _
	"WHNT_News-Top.xml" & vbTAB & "http://feeds2.feedburner.com/whnt-news" & vbTAB & "www.WHNT.com", _
	"WHNT_News-Hsv.xml" & vbTAB & "http://feeds2.feedburner.com/whnt-huntsville-madison-county-news" & vbTAB & "www.WHNT.com", _
	"WAAY_News.xml" & vbTAB & "http://feeds.feedburner.com/Waay-WaayNewsAndHome" & vbTAB & "www.WAAY.com", _
	"WAAY_News-Local.xml" & vbTAB & "http://feeds.feedburner.com/Waay-LocalNews" & vbTAB & "www.WAAY.com", _
	"WAAY_News-State.xml" & vbTAB & "http://feeds.feedburner.com/Waay-StateNews" & vbTAB & "www.WAAY.com", _
	"WAFF_News.xml" & vbTAB & "http://www.waff.com/global/Category.asp?c=14421&clienttype=rss" & vbTAB & "www.WAFF.com", _
	"-", _
	"Fox-News.xml" & vbTAB & "http://feeds.foxnews.com/foxnews/latest", _
	"Fox-National.xml" & vbTAB & "http://feeds.foxnews.com/foxnews/national", _
	"AP-Top_News.xml" & vbTAB & "http://hosted.ap.org/lineups/TOPHEADS-rss_2.0.xml?SITE=KFWB&SECTION=HOME", _
	"AP-US_News.xml" & vbTAB & "http://hosted.ap.org/lineups/USHEADS-rss_2.0.xml?SITE=KYB66&SECTION=HOME", _
	"UPI-Top_News.xml" & vbTAB & "http://www.upi.com/rss/Top_News/", _
	"Yahoo-Top_Stories.xml" & vbTAB & "http://rss.news.yahoo.com/rss/topstories", _
	"-", _
	"Yahoo-Weather_Detailed.xml" & vbTAB & "http://xml.weather.yahoo.com/forecastrss?p=USAL0287&u=f", _
	"Yahoo-Weather.xml" & vbTAB & "http://xml.weather.yahoo.com/forecastrss?p=35801" _
	)

'	"MC_Events_AL.xml" & vbTAB & "feed://www.weaselsspfld.com/rss/motorcycle/alabama.xml" & vbTAB & "www.weaselsspfld.com", _
'	"MC_Events_GA.xml" & vbTAB & "http://www.weaselsspfld.com/rss/motorcycle/georgia.xml" & vbTAB & "www.weaselsspfld.com", _
'	"MC_Events_MS.xml" & vbTAB & "http://www.weaselsspfld.com/rss/motorcycle/mississippi.xml" & vbTAB & "www.weaselsspfld.com", _
'	"MC_Events_TN.xml" & vbTAB & "http://www.weaselsspfld.com/rss/motorcycle/tennessee.xml" & vbTAB & "www.weaselsspfld.com", _

'	"Microsoft_Watch.xml" & vbTAB & "http://rssnewsapps.ziffdavis.com/msw.xml", _
'	"WindowsITPro_SQL.xml" & vbTAB & "http://feeds.penton.com/windowsitpro/Ptva", _
'	"WindowsITPro_WinAdmin.xml" & vbTAB & "http://feeds.penton.com/windowsitpro/Xluj", _
'	"UnitedMethodistNews.xml" & vbTAB & "http://archives.umc.org/rss/RSS_UMNS15.xml", _
'	"wunderground.xml" & vbTAB & "http://rss.wunderground.com/auto/rss_full/AL/Huntsville.xml?units=english" _
'	"CcomBest.xml" & vbTAB & "http://media.salemwebnetwork.com/RSS/CCOM/todaysbestoftheweb.xml", _
'	"CcomDevotion.xml" & vbTAB & "http://media.salemwebnetwork.com/RSS/CCOM/dailydevotional.xml", _
'	"CcomDailyVerse.xml" & vbTAB & "http://media.salemwebnetwork.com/RSS/CCOM/verseoftheday.xml", _
'	"USRiderNews.xml" & vbTAB & "http://www.usridernews.com/absolutenm/rss.asp", _
'	"relevantmagazine.xml" & vbTAB & "http://www.relevantmagazine.com/rss/relevantmagazine.xml", _
'	"internettop.xml" & vbTAB & "http://www.internetnews.com/icom_includes/feeds/inews/xml_front-10.xml", _
'	"yahoomostviewed.xml" & vbTAB & "http://rss.news.yahoo.com/rss/mostviewed", _
'	"foxus.xml" & vbTAB & "http://www.foxnews.com/xmlfeed/rss/0,4313,1,00.rss", _
'	"msdnfrontpage.xml" & vbTAB & "http://msdn.microsoft.com/office/frontpage/rss.xml", _
'	"GoogleXPS.xml" & vbTAB & "http://news.google.com/news?hl=en&ned=us&q=xml+paper+specification&ie=UTF-8&output=rss", _
'	"GoogleKnowledge.xml" & vbTAB & "http://news.google.com/news?hl=en&ned=us&ie=UTF-8&q=Knowledge+Management&output=rss", _
'	"GoogleVista.xml" & vbTAB & "http://news.google.com/news?hl=en&ned=us&ie=UTF-8&q=Windows+Vista&output=rss", _

nNewsIndex = Request("rss")
IF "" <> nNewsIndex THEN
	IF ISNUMERIC(nNewsIndex) THEN
		nNewsIndex = CINT(nNewsIndex)
	ELSE
		nNewsIndex = 0
	END IF
	IF UBOUND(aNews) < nNewsIndex THEN nNewsIndex = 0
ELSE
	nNewsIndex = Request.Cookies("personal_news")
	IF "" = nNewsIndex THEN
		nNewsIndex = 0
	ELSE
		nNewsIndex = CINT(nNewsIndex) '+ 1
		IF UBOUND(aNews) < nNewsIndex THEN nNewsIndex = 0
		IF "-" = aNews(nNewsIndex) THEN nNewsIndex = nNewsIndex + 1
	END IF
END IF
Response.Cookies("personal_news") = nNewsIndex
Response.Cookies("personal_news").Expires = DateAdd( "d", 14, NOW )
	





%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\include_theme.asp"-->

<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/remind_cache.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!--#include file="scripts\rss.asp"-->
<!--#include file="scripts\rssweather.asp"-->
<%





SUB loadCachedFile( sCache, sFile )

	DIM	sPath
	sPath = cache_filepath( sCache, sFile )
	IF g_oFSO.FileExists( sPath ) THEN
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sPath, 1 )
		IF NOT Nothing IS oFile THEN
			DIM	sContent
			sContent = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			Response.Write sContent
		END IF
	ELSE
		Response.Write "Loading ..."
	END IF

END SUB


FUNCTION loadCachedFileX( sCache, sFile, sInterval, nIntValue, sBreakInterval )
	loadCachedFileX = TRUE

	DIM	sPath
	DIM	oFile
	DIM	sContent
	sPath = cache_checkFile( sCache, sFile, sInterval, nIntValue, sBreakInterval )
	IF 0 < LEN(sPath) THEN
		SET oFile = g_oFSO.OpenTextFile( sPath, 1 )
		IF NOT Nothing IS oFile THEN
			sContent = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			Response.Write sContent
			loadCachedFileX = FALSE
		END IF
	ELSE
		sPath = cache_filepath( sCache, sFile )
		IF g_oFSO.FileExists( sPath ) THEN
			SET oFile = g_oFSO.OpenTextFile( sPath, 1 )
			IF NOT Nothing IS oFile THEN
				sContent = oFile.ReadAll
				oFile.Close
				SET oFile = Nothing
				Response.Write sContent
			END IF
		ELSE
			Response.Write "Loading ..."
		END IF
	END IF

END FUNCTION




%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<title>Home</title>
<meta name="navigate" content="tab">
<meta name="navtitle" content="Home">
<meta name="sortname" content="aaaaadefault">
<meta content="Microsoft FrontPage 12.0" name="GENERATOR">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<link rel="stylesheet" href="<%=remindCSS()%>" type="text/css">
<style type="text/css">

.rssTitle
{
	font-weight: bold;
	font-family: sans-serif;
	font-size: large;
}

div.rssItem
{
	margin-top: 1em;
	position: relative;
	display: block;
	clear: both;
}

.rssItem p
{
	margin-top: 0;
	margin-bottom: 0;
}

.rssDate
{
	color: #999999;
	font-size: smaller;
	font-family: sans-serif;
	clear: both;
}


#rsscontent
{
	position: relative;
}

.remind_table
{
	width: 100%;
}





</style>
<link rel="stylesheet" href="scripts/rssweather.css" type="text/css">
<script language="JavaScript" type="text/javascript">
<!--

function createEditWindow()
{
	var w, h;
	var x, y;
	var s;

	//	First calculate the size and position for the window
	w = 800;
	h = 800;
	
	//	We need to test if the screen object exists because
	//	older browsers don't support it.
	if ( window.screen )
	{
		x = (screen.width-w)/2;
		y = (screen.height-h)/3;
	}
	else
	{
		x = 100;
		y = 100;
	}
	s = "width=" + w + ",height=" + h + ",left=" + x + ",top=" + y;
	
	//	create the window with just a title bar
	window.open( "maintenance/remind.asp",
					"CalendarEditWindow",
					s + ",resizable=yes,scrollbars=yes,status=no" );
}

//-->
</script>
<script type="text/javascript">
<!--

function doLoad()
{
	var h;
	var w;
	
	//if ( screen.availHeight < screen.height )
	//	h = screen.availHeight - 15;
	//else
	//	h = screen.height - 200;
	//window.resizeTo( 840, h );
	//top.window.resizeTo( 1000, h );
	//top.window.resizable
	
	/*
	w = top.window;
	if ( w.screenTop  &&  w.screenLeft )
	{
		if ( 10 < w.screenTop )
		{
			//if ( document.body.scrollLeft )
				w.moveTo( document.body.scrollLeft + w.screenLeft, w.availTop+5 );
			//else
			//	w.moveTo( w.screenLeft, 5 );
		}
	}
	else if ( w.screenY )
	{
		if ( w.availTop+19 < w.screenY )
			w.moveTo( document.body.scrollLeft + w.screenX, w.availTop+5 );
	}
	*/
}





var loadedobjects=""
var rootdomain="http://"+window.location.host

function ajaxPage(url, containerid)
{
	var page_request = false;
	if (window.XMLHttpRequest) // if Mozilla, Safari etc
	{
		page_request = new XMLHttpRequest();
	}
	else if (window.ActiveXObject) // if IE
	{
		try
		{
			page_request = new ActiveXObject("Msxml2.XMLHTTP");
		}
		catch (e)
		{
			try
			{
				page_request = new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (e)
			{
				return false;
			}
		}
	}
	else
	{
		return false;
	}
	
	var sURL = url;
	var sUID = "zzz=" + escape((new Date()).toUTCString());
	sURL += (sURL.indexOf("?")+1) ? "&" : "?";
	sURL += sUID;

	page_request.onreadystatechange=function(){
		ajaxHandleStateChange(page_request, containerid);
	}
	
	page_request.sTargetURL = url;
	window.status = "Request to Load = " + url;
	page_request.open('GET', sURL, true);
	page_request.send(null);
}

function ajaxHandleStateChange(page_request, containerid)
{
	if (page_request.readyState == 4 )
	{
		if ( page_request.status == 200
				||	window.location.href.indexOf("http") == -1 )
		{
			document.getElementById(containerid).innerHTML=page_request.responseText;
			window.status = "loaded external url";
			setTimeout( "ajaxClearStatus()", 3*1000 );
		}
		else
		{
			window.status = "retry loading external url";
			page_request.open( "GET", page_request.sTargetURL, true );
			page_request.send( null );
		}
	}
}

function ajaxClearStatus()
{
	window.status = "";
}

function loadobjs()
{
	if (!document.getElementById)
		return;
	for (i=0; i<arguments.length; i++)
	{
		var file=arguments[i];
		var fileref="";
		if (loadedobjects.indexOf(file)==-1) //Check to see if this object has not already been added to page before proceeding
		{
			if (file.indexOf(".js")!=-1) //If object is a js file
			{
				fileref=document.createElement('script');
				fileref.setAttribute("type","text/javascript");
				fileref.setAttribute("src", file);
			}
			else if (file.indexOf(".css")!=-1) //If object is a css file
			{
				fileref=document.createElement("link");
				fileref.setAttribute("rel", "stylesheet");
				fileref.setAttribute("type", "text/css");
				fileref.setAttribute("href", file);
			}
		}
		if (fileref!="")
		{
			document.getElementsByTagName("head").item(0).appendChild(fileref);
			loadedobjects+=file+" "; //Remember this object as being already added to page
		}
	}
}






//-->
</script>
<style type="text/css">
<!--
.line        { border-bottom: 2px solid black }
.headline    { font-family: sans-serif; font-size: xx-small }
.title    { font-family: sans-serif; font-size: xx-small }
.RemindMonthHeader { background-image: url('images/bkg_remindMonthHeader.jpg') }

.SmileBody .htmlformattable
{
	color: black;
	background-color: gold;
	font-size: small;
	font-family: Arial, sans-serif;
	border-color: black;
	text-align: center;
}

.SmileBody .htmlformattable td
{
	border-color: black;
}

.SmileBody h4.htmlformat
{
	color: #CC9966;
	border-bottom: 2px solid #CC9966;
}



.SmileBody h3.htmlformat, .SmileBody h4.htmlformat, .SmileBody h5.htmlformat
{
	font-family: sans-serif;
}

.SmileBody h5.htmlformat
{
	color: #996600;
}






-->
</style>

<%
makeWeatherStyles
%>


</head>

<body bgcolor="#FFFFFF" onload="doLoad()">
<!--#include file="scripts\include_navigation.asp"-->
<%
'outputPad2
%>

<table border="0" cellpadding="0" cellspacing="4">
  <tr>
    <td valign="top">


<table border="0" cellspacing="2" width="225" cellpadding="0">
  <tr>
    <th align="left" class="line" bgcolor="#99CCFF">Weather</th>
  </tr>
  <tr>
    <td>
          <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
          <td bgcolor="#000000">
          <!--webbot bot="Include"
            u-include="_private/moonphas.htm" tag="BODY" startspan -->

<table border="0" cellpadding="3" cellspacing="0">
  <tr>
    <td valign="middle" align="center" bgcolor="#000000" width="1"><img
    src="images/moon/lunardot.gif" id="LunarDot" name="LunarDot" WIDTH="1" HEIGHT="1"></td>
    <td valign="middle" align="center" bgcolor="#000000"><script language="JavaScript">
<!--

var gnMoonAge = 0;
var gsMoonPhase = "";
var gsMoonString = "";
var goDateNow = new Date();
var kLunarCycle = 29.530588853;
var kTimezoneOffset = goDateNow.getTimezoneOffset() * 60 * 1000;
var kTimezoneOffsetDay = goDateNow.getTimezoneOffset() / 60 / 24;
var goDateNowUTC = new Date(goDateNow.getTime() + kTimezoneOffset);

function calcMoonAge( year, month, day, hour, minutes )
{
	var YY = year;
	var MM = month;
	var DD = day;
	var HR = hour;
	var MN = minutes;
	var GGG = 1;
	var JD = 0;
	var S = 1;
	var A = 0;
	var J1 = 0;
	var J = 0;
	var V = 0;
	var IP = 0;

	with ( Math )
	{  
		HR = HR + (MN / 60);
		GGG = 1;
		if ( YY <= 1585 )
			GGG = 0;
		JD = -1 * floor(7 * (floor((MM + 9) / 12) + YY) / 4);
		S = 1;
		if ( (MM - 9) < 0 )
			S=-1;
		A = abs(MM - 9);
		J1 = floor(YY + S * floor(A / 7));
		J1 = -1 * floor((floor(J1 / 100) + 1) * 3 / 4);
		JD = JD + floor(275 * MM / 9) + DD + (GGG * J1);
		JD = JD + 1721027 + 2 * GGG + 367 * YY - 0.5;
		J = JD + (HR / 24);
	}
		
	/* Calculate Illumination (synodic) phase */
	
	/*P2 = 2 * 3.141592654;*/
	V = (J-2451550.1) / kLunarCycle;
	IP = normalizeValue( V );
	return IP * kLunarCycle;
}

function normalizeValue( v )
{
	var nValue = v - Math.floor(v); 
	if ( nValue < 0 )
		nValue = nValue + 1;
	return nValue;
}


function getMoonPhase( nAge )
{
	if ( nAge < 0.5 ) return "New Moon";
	if ( nAge < 1.5 ) return "Past New";
	if ( nAge < kLunarCycle*0.25 - 1 ) return "Waxing Cresent";
	if ( nAge < kLunarCycle*0.25 + 1 ) return "First Quarter";
	if ( nAge < kLunarCycle*0.5 - 1.5 ) return "Waxing Gibbous";
	if ( nAge < kLunarCycle*0.5 - .5 ) return "Near Full";
	if ( nAge < kLunarCycle*0.5 + .5 ) return "Full Moon";
	if ( nAge < kLunarCycle*0.5 + 1.5 ) return "Past Full";
	if ( nAge < kLunarCycle*0.75 - 1 ) return "Waning Gibbous";
	if ( nAge < kLunarCycle*0.75 + 1 ) return "Last Quarter";
	if ( nAge < kLunarCycle - 1.5 ) return "Waning Cresent";
	if ( nAge < kLunarCycle - .5 ) return "Near New";
	return "New Moon";
}

monthNames = new Array(13);
monthNames[1]  = "Jan";
monthNames[2]  = "Feb";
monthNames[3]  = "Mar";
monthNames[4]  = "Apr";
monthNames[5]  = "May";
monthNames[6]  = "Jun";
monthNames[7]  = "Jul";
monthNames[8]  = "Aug";
monthNames[9]  = "Sep";
monthNames[10] = "Oct";
monthNames[11] = "Nov";
monthNames[12] = "Dec";

function getLongDate( dateObj )
{
	var theMonth = monthNames[dateObj.getMonth()+1];
	var theDate = dateObj.getDate();
	var theYear = dateObj.getYear();
	var sDate = "";
	var sHours = "";
	var sMins = "";
	if ( theYear < 1000 )
		theYear += 1900;
	if ( theDate < 10 )
		sDate = "0" + theDate;
	else
		sDate = "" + theDate;
	var theHours = dateObj.getHours();
	if ( theHours < 10 )
		sHours = "0" + theHours;
	else
		sHours = "" + theHours
	var theMinutes = dateObj.getMinutes();
	if ( theMinutes < 10 )
		sMins = "0" + theMinutes;
	else
		sMins = "" + theMinutes;
	return "" + sDate + "-" + theMonth + "-" + theYear + " " + sHours + ":" + sMins;
}
		
function getNextFull( moonAge )
{	
	var currMilSecs = goDateNowUTC.getTime();
	var daysToGo = kLunarCycle/2.0 - moonAge;
	while( daysToGo < 0.5 )
	{
		daysToGo = daysToGo + kLunarCycle;
	}
	var milSecsToGo = daysToGo*24*60*60*1000;
	var nextMoonTime = currMilSecs + milSecsToGo;
	var nextMoonDate = new Date(nextMoonTime);
	return nextMoonDate;
}

function getNextNew( moonAge )
{	
	var currMilSecs = goDateNowUTC.getTime();
	var daysToGo = kLunarCycle - moonAge;
	while( daysToGo < 0.5 )
	{
		daysToGo = daysToGo + kLunarCycle;
	}
	var milSecsToGo = daysToGo*24*60*60*1000;
	var nextMoonTime = currMilSecs + milSecsToGo;
	var nextMoonDate = new Date(nextMoonTime);
	return nextMoonDate;
}

function writeMoonGIF()
{
	var thePict;
	var n = 0;
	var sn = "";

	if ( document.all )
		thePict = document.all.LunarDot;
	else if ( document.images )
		thePict = document.images["LunarDot"];
	else
		return;

	n = Math.ceil(gnMoonAge - 0.5);
	if ( 29.25 < gnMoonAge )
		n = 0;
	else if ( n < 0 )
		n = 0;
	if ( n < 10 )
		sn = "0" + n;
	else
		sn = "" + n;
	var thePictSrc = thePict.src;
	var sTemp = thePictSrc.substring( 0, thePictSrc.indexOf("/images"));
	document.write('<a href="' + sTemp + '/moon.htm" target="_self">');
	document.write('<img src="' + sTemp + '/images/moon/mm' + sn + '.gif" alt="Moon Phase" height="40" width="40" border="0">');
	document.write('</a>');
}


	var theYear = goDateNowUTC.getYear();
	var theMonth = goDateNowUTC.getMonth()+1;
	var theDay = goDateNowUTC.getDate();
	var theHour = goDateNowUTC.getHours();
	var theMinutes = goDateNowUTC.getMinutes();

	if ( theYear < 1000 )
		theYear += 1900;

	gnMoonAge = calcMoonAge( theYear, theMonth, theDay, theHour, theMinutes );
	gsMoonPhase = getMoonPhase( gnMoonAge );
	if ( 29.53 < gnMoonAge)
		gnMoonAge = 0;

	gsMoonString = '<table border="0" cellspacing="0" cellpadding="0">';
	gsMoonString += '<tr><td colspan="2">';
	gsMoonString += '<font face="Arial,Helvetica" color="#FFFFFF"><b>' + gsMoonPhase + '</b></font>';
	gsMoonString += '<td></tr>';
	/*gsMoonString += "Lunar age: " + gnMoonAge + "<br>";*/

	var dateNew = getNextNew( gnMoonAge );
	var dateFull = getNextFull( gnMoonAge );

	if ( dateNew < dateFull )
	{
		gsMoonString += '<tr>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">Next New:&nbsp;</font></td>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">' + getLongDate(dateNew) + '</font></td>';
		gsMoonString += '</tr><tr>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">Next Full:&nbsp;</font></td>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">' + getLongDate(dateFull) + '</font></td>';
		gsMoonString += '</tr>';
	}
	else
	{
		gsMoonString += '<tr>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">Next Full:&nbsp;</font></td>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">' + getLongDate(dateFull) + '</font></td>';
		gsMoonString += '</tr><tr>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">Next New:&nbsp;</font></td>';
		gsMoonString += '<td><font face="Arial,Helvetica" color="#FFFFFF" size="1">' + getLongDate(dateNew) + '</font></td>';
		gsMoonString += '</tr>';
	}
	gsMoonString += "</table>";


	writeMoonGIF();
//-->
    </script></td>
    <td bgcolor="#000000" style="font-family: sans-serif"><script type="text/javascript" language="JavaScript"><!--
		document.write(gsMoonString);
//--></script></td>
  </tr>
</table>
<!--webbot bot="Include" endspan i-checksum="19032" --></td>
        </tr>
      </table>
<%
DIM	ZIPWeather
ZIPWeather = 35801
%>

<script type="text/javascript">

function rssloadweatherajax()
{
	ajaxPage('loadrssweather.asp?z=<%=ZIPWeather%>', 'rssweathercontent');
}
function timerloadweatherajax()
{
	setTimeout( "rssloadweatherajax()", 0.25*1000 );
}

if (window.addEventListener)
	window.addEventListener("load", timerloadweatherajax, false)
else if (window.attachEvent)
	window.attachEvent("onload", timerloadweatherajax)
else
	setTimeout( "timerloadweatherajax()", 2.0*1000 );



</script>
<div id="rssweathercontent">
<%

IF loadCachedFileX( "rss", "AWSWxCurrent" & ZIPWeather & ".htm", "n", 20, "h" ) THEN
%>
<script type="text/javascript">
if (window.addEventListener)
	window.addEventListener("load", timerloadweatherajax, false)
else if (window.attachEvent)
	window.attachEvent("onload", timerloadweatherajax)
else
	setTimeout( "timerloadweatherajax()", 2.0*1000 );
</script>
<%
END IF

%>
</div>

<ul>
    <!--li><a href="http://www.wunderground.com/US/AL/Madison.html">
<img src="http://banners.wunderground.com/banner/infobox_both/language/www/US/AL/Madison.gif"
alt="Click for Madison, Alabama Forecast" height=108 width=144 border="0" nosave></a></li-->
        <li><a
          href="../grizzlyweb/links/weather_current.asp?loc=al_huntsville">Current
          Conditions</a></li>
        <li><a
          href="../grizzlyweb/links/weather_forecast.asp?loc=al_huntsville">Forecast</a></li>
        <li><a
          href="../grizzlyweb/links/weather_forecast_maps.asp?loc=al_huntsville">Forecast
          Maps</a></li>
        <li><a href="../grizzlyweb/links/weather_radar.asp?loc=al_huntsville">Local
          Radar</a></li>
        <li><a
          href="../grizzlyweb/links/weather_radar_region.asp?loc=al_huntsville">Regional
          Radar</a></li>
        <li><a
          href="../grizzlyweb/links/weather_warnings.asp?loc=al_huntsville">Storm
          Warnings</a></li>
        <li><a href="../grizzlyweb/links/weather.asp?loc=al_huntsville">Weather
          Links</a></li>
      </ul>
    </td>
  </tr>
  <tr>
    <td align="center">
<script type="text/javascript">

function loadhumorajax()
{
	ajaxPage('loadhumor.asp', 'rsshumorcontent');
}
function timerloadhumorajax()
{
	setTimeout( "loadhumorajax()", 0.4*1000 );
}

if (window.addEventListener)
	window.addEventListener("load", timerloadhumorajax, false)
else if (window.attachEvent)
	window.attachEvent("onload", timerloadhumorajax)
else
	setTimeout( "timerloadhumorajax()", 2.0*1000 );



</script>
<div id="rsshumorcontent">Loading Humor ...</div>


    </td>
  </tr>
</table>
      </td>
      <td width="2" bgcolor="#CCCCCC"><spacer type="block" height="1" width="1"></td>
    <td valign="top">
<%

DIM	nDateBegin
DIM	nDateEnd

nDateBegin = jdateFromVBDate( NOW )
nDateEnd = nDateBegin + 28 - 1
nDateEnd = -21	'pending



%>

<script type="text/javascript">

function loadcalendarajax()
{
	ajaxPage("loadcalendar.asp?f=homeevents.htm&i=h&v=8&b=d&db=" + "<%=nDateBegin%>" +"&de="+"<%=nDateEnd%>"+"&h="+escape('drs,none')+"&s=true&x=" +escape('scripts/remind.xslt'), 'rsscalendarcontent');
}
function timerloadcalendarajax()
{
	setTimeout( "loadcalendarajax()", 0.3*1000 );
}




</script>
<table border="0" width="350" cellspacing="0" cellpadding="2">
  <tr>
    <td width="100%" bgcolor="#FFDD99" class="line">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <th align="left">Calendar</th>
          <td align="right"><a href="javascript:createEditWindow()">edit...</a>
            &nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
  <td width="100%">

<div id="rsscalendarcontent">
<%
IF loadCachedFileX( "remind", "homeevents.htm", "h", 8, "d" ) THEN
%>
<script type="text/javascript">
if (window.addEventListener)
	window.addEventListener("load", timerloadcalendarajax, false)
else if (window.attachEvent)
	window.attachEvent("onload", timerloadcalendarajax)
else
	setTimeout( "timerloadcalendarajax()", 2.0*1000 );
</script>
<%
END IF
%>
</div>
</td>
</tr>
</table>
<%

'===========================================================




%>
<hr>
	  <a href="maintenance/clearcache.asp" target="_blank">Clear Cache</a>
    </td>
      <td width="2" bgcolor="#CCCCCC"><spacer type="block" height="1" width="1"></td>
    <td valign="top">
    <a href="../grizzlyweb/links/index.asp?loc=al_huntsville">
	<img border="0" src="bearpawc3.gif"  width="40" height="50"></a>&nbsp;&nbsp;&nbsp;
	<a href="http://www.RocketCityWings.org/" target="_blank">
	<img src="../RocketCityWings/favicon48.gif" border="0" height="48" width="48" alt="Rocket City Wings" ></a>&nbsp;&nbsp;
	<a href="http://www.Alabama-gwrra.org/" target="_blank"> 
	<img src="../Alabama-gwrra/favicon48.gif" border="0" height="48" width="48" alt="Alabama-gwrra" ></a>&nbsp;&nbsp;&nbsp; <a href="http://www.facebook.com/" target="_blank">
	<img alt="Facebook" src="images/f_logo.jpg" border="0" height="48" width="48"></a>&nbsp; &nbsp;
<a href="http://www.Linkedin.com" target="_blank">
	<img alt="LinkedIn" src="images/LinkedIn_IN_Icon_45px.png" border="0" height="45" width="45"></a>&nbsp;&nbsp;
<a href="http://www.indeed.com" target="_blank">
	<img alt="LinkedIn" src="images/indeed_i_bigger.png" border="0" height="45" width="45"></a>
    <!-- Web search from Bing-->
<form method="get" action="http://www.bing.com/search">
<input type="hidden" name="cp" value="CODE PAGE USED BY YOUR HTML PAGE" />
<input type="hidden" name="FORM" value="FREEWS" />
  <table bgcolor="#FFFFFF">
    <tr>
      <td>
        <a href="http://www.bing.com/">
          <img src="images/bing-favicon.jpg" border="0" ALT="bing" /></a>&nbsp;
      </td>
      <td>
        <input type="text" name="q" size="30" />
        <input type="submit" value="Search" />
      </td>
    </tr>
  </table>
</form>
<!-- Web Search from Bing -->



<FORM method=GET action="http://www.google.com/search">
<input type=hidden name=ie value=UTF-8>
<input type=hidden name=oe value=UTF-8>
<TABLE bgcolor="#FFFFFF"><tr><td>
<A HREF="http://www.google.com/">
<IMG SRC="images/google-favicon.png"
border="0" ALT="Google" align="absmiddle" height="32" width="32"></A>
<INPUT TYPE=text name=q size=30 maxlength=255 value="">
<INPUT type=submit name=btnG VALUE="Search">
</td></tr></TABLE>
</FORM>

<br>
    <form method="get" action="default.asp" name="rssform" style="margin-bottom: 0">
<script type="text/javascript" language="javascript">

function rsschange( o )
{
	var oItems;
	var oForm;
	oItems = document.getElementsByName("rssform");
	if ( oItems )
	{
		oForm = oItems.item(0);
		if ( oForm )
			oForm.submit();
	}
}





</script>
    <select name="rss" onchange="rsschange(this)">
    <%
SUB makeRSSOptions
	DIM	j
	DIM	i
	DIM	aSplit
	DIM	sName
	FOR i = 0 TO UBOUND(aNews)
		IF "" <> aNews(i) THEN
			IF "-" = aNews(i) THEN
				Response.Write "<option value=""" & i+1 & """>------------</option>" & vbCRLF
			ELSE
				aSplit = SPLIT(aNews(i),vbTAB)
				Response.Write "<option value=""" & i & """"
				IF i = nNewsIndex THEN Response.Write " selected"
				Response.Write ">"
				j = INSTR(aSplit(0),".")
				sName = LEFT(aSplit(0),j-1)
				sName = REPLACE( sName, "_", " " )
				Response.Write Server.HTMLEncode(sName)
				Response.Write "</option>" & vbCRLF
			END IF
		END IF
	NEXT 'i
END SUB
makeRSSOptions
    %>
    </select>

</form>
<%

SUB AMADirectLink

	ON ERROR Goto 0
	
	
	RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
	
	DIM	j
	j = INT( (UBOUND(aNews)+1) * RND + 0.5 ) - 1
	j = nNewsIndex
	IF UBOUND(aNews) < j THEN
		j = UBOUND(aNews)
	ELSEIF j < 0 THEN
		j = 0
	END IF
	
	IF "-" = aNews(j) THEN
		j = j + 1
	END IF
	
	DIM	aSplitRSS
	aSplitRSS = SPLIT( aNews(j), vbTAB )

	DIM	sFile
	DIM	sFileEncode
	DIM	sURL
	DIM	sFail

	sFile = aSplitRSS(0)
	sFileEncode = Server.URLEncode(sFile)
	sURL = Server.URLEncode(aSplitRSS(1))
	IF 1 < UBOUND(aSplitRSS) THEN
		sFail = Server.URLEncode(aSplitRSS(2))
	ELSE
		sFail = ""
	END IF

%>
<script type="text/javascript">

function rssloadajax()
{
	ajaxPage('loadrssnews.asp?f=<%=sFileEncode%>&h=<%=sURL%>&x=<%=sFail%>', 'rsscontent');
}
function timerloadajax()
{
	setTimeout( "rssloadajax()", 0.25*1000 );
}


//setTimeout( "rssloadajax()", 0.25*1000 );

</script>
<div id="rsscontent">
<%
IF loadCachedFileX( "rss", sFile & ".htm", "h", 4, "d" ) THEN
%>
<script type="text/javascript">
if (window.addEventListener)
	window.addEventListener("load", timerloadajax, false)
else if (window.attachEvent)
	window.attachEvent("onload", timerloadajax)
else
	setTimeout( "timerloadajax()", 2.0*1000 );
</script>
<%
END IF
%>
</div>
<%
	
	
END SUB
AMADirectLink


%>
    </td>
  </tr>
</table>
<!--webbot bot="Include" u-include="_private/byline.htm" tag="BODY" startspan -->

<script language="JavaScript">
<!--

function makeByLine()
{
	document.write( '<' + 'script language="JavaScript" src="http://' );
	if ( "localhost" == location.hostname )
	{
		document.writeln( 'localhost/BearConsultingGroup' );
	}
	else
	{
		document.writeln( 'BearConsultingGroup.com' );
	}
	document.writeln( '/designby_small.js"></' + 'script>' );
}
makeByLine()

//-->
</script>
<!--script language="JavaScript" src="http://BearConsultingGroup.com/designbyadvert.js"></script-->

<!--webbot bot="Include" endspan i-checksum="20551" -->


</body>

</html>
<%
SET g_oFSO = Nothing
%>