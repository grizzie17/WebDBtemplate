<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT



%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\include_pagelist.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/remind_cache.asp"-->
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts\include_gallery.asp"-->
<!--#include file="scripts/include_calendar.asp"-->
<!--#include file="scripts\include_announce.asp"-->
<!--#include file="scripts\include_rssannounce.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\rss.asp"-->
<!--#include file="scripts\include_forum.asp"-->
<!--#include file="scripts\rssweather.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_emailnotificationformat.asp"-->
<!--#include file="scripts/include_pagebody.asp"-->
<!--#include file="scripts/include_picture2.asp"-->
<!--#include file="scripts/index_tools.asp"-->
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





DIM g_n
DIM g_sGalleryFilename
DIM g_aGalleryFilenames(29)
DIM	g_sGalleryFolder
DIM g_nGalleryIndexCount
DIM g_dGalleryIndexTime
DIM g_sGalleryCookie
g_sGalleryFilename = ""
g_sGalleryFolder = ""
g_sGalleryCookie = ""
g_nGalleryIndexCount = 0
g_dGalleryIndexTime = CDATE("1/1/2000")
FOR g_n = LBOUND(g_aGalleryFilenames) TO UBOUND(g_aGalleryFilenames)
	g_aGalleryFilenames(g_n) = ""
NEXT


DIM g_sHumorFilename
DIM	g_sHumorFolder
DIM	g_nHumorIndexCount
DIM g_dHumorIndexTime
DIM g_sHumorCookie
g_sHumorFilename = ""
g_sHumorFolder = ""
g_sHumorCookie = ""
g_nHumorIndexCount = 0
g_dHumorIndexTime = CDATE("1/1/2000")


DIM g_sSafetyFilename
DIM g_sSafetyFolder
DIM g_nSafetyIndexCount
DIM g_dSafetyIndexTime
DIM g_sSafetyCookie
g_sSafetyFilename = ""
g_sSafetyFolder = ""
g_sSafetyCookie = ""
g_nSafetyIndexCount = 0
g_dSafetyIndexTime = CDATE("1/1/2000")



FUNCTION findShuffleIndex( sBaseName )

	findShuffleIndex = ""

	DIM	sFolder

	sFolder = findAppDataFolder( sBaseName )
	IF "" <> sFolder THEN
	
		EXECUTEGLOBAL "g_s" & sBaseName & "Folder = """ & sFolder & """"

		DIM	sIndexName
		sIndexName = cache_filepath( "shuffle", sBaseName&".txt" )
		'sIndexName = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
		
		IF g_oFSO.FileExists( sIndexName ) THEN
			findShuffleIndex = sIndexName
		END IF
	
	END IF

END FUNCTION


FUNCTION findHumorIndex()
	findHumorIndex = findShuffleIndex( "Humor" )
END FUNCTION


FUNCTION findSafetyIndex()
	findSafetyIndex = findShuffleIndex( "Safety" )
END FUNCTION


FUNCTION getNextCookie( sBaseName, nLow, nHigh )

	DIM	nCookie
	DIM sCookie
	DIM	sCookieName
	sCookieName = g_sCookiePrefix & "_" & sBaseName & "_index"
	sCookie = Request.Cookies( sCookieName )
	IF "" <> sCookie  AND  ISNUMERIC(sCookie) THEN
		nCookie = CINT(sCookie)
		nCookie = nCookie + 1
		IF nHigh < nCookie THEN
			nCookie = nLow
		ELSEIF nCookie < nLow THEN
			nCookie = nLow
		END IF
	ELSE
		DIM j
		RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
		nCookie = ROUND( RND * ( nHigh - nLow )) + nLow
		IF nHigh < nCookie THEN
			nCookie = nHigh
		ELSEIF nCookie < nLow THEN
			nCookie = nLow
		END IF
	END IF
	Response.Cookies( sCookieName ) = nCookie
	Response.Cookies( sCookieName ).Expires = DateAdd( "d", 56, NOW )
	
	getNextCookie = nCookie
	
	EXECUTEGLOBAL "g_s" & sBaseName & "Cookie = " & nCookie

END FUNCTION


SUB readTextFolderObject( sRoot, oInFolder, aLetters(), nCountLetters )

	IF oInFolder IS Nothing THEN
		EXIT SUB
	END IF
	IF "_" = LEFT(oInFolder.Name,1) THEN
		EXIT SUB
	END IF
	
	DIM	oFolder
	DIM	oFile
	DIM	sFile
	DIM	i
	DIM	nLenRoot
	
	nLenRoot = LEN(sRoot)+1

	SET oFolder = oInFolder
			
	i = nCountLetters
	FOR EACH oFile IN oFolder.Files
		sFile = LCASE(oFile.Name)
		IF "_" <> LEFT(sFile,1) THEN
			IF 0 < INSTR(1, sFile, ".txt", vbTextCompare ) THEN
				IF UBOUND(aLetters) <= i THEN
					REDIM PRESERVE aLetters(i+50)
				END IF
				aLetters(i) = MID(oFile.Path,nLenRoot)
				i = i + 1
			END IF
		END IF
	NEXT 'oFile
	SET oFile = Nothing
	
	DIM oSubFolder
	FOR EACH oSubFolder IN oFolder.SubFolders
		IF "_" <> LEFT(oSubFolder.Name,1)  AND  "images" <> LCASE(oSubFolder.Name) THEN
			readTextFolderObject sRoot, oSubFolder, aLetters, i
		END IF
	NEXT 'oSubFolder
	
	nCountLetters = i
	


END SUB


SUB readTextListFolder( sFolder, aLetters() )

	
	IF "" <> sFolder THEN
		DIM	oFolder
		DIM	oFile
		DIM	sFile
		DIM sRoot
		DIM	i
		
		SET oFolder = g_oFSO.GetFolder( sFolder )
		
		sRoot = sFolder
		IF "\" <> RIGHT(sRoot,1) THEN
			sRoot = sRoot & "\"
		END IF
		
		i = 0
		REDIM aLetters(10)
		readTextFolderObject sRoot, oFolder, aLetters, i
		
		IF 0 < i THEN
			REDIM PRESERVE aLetters(i-1)
			shuffle aLetters, 0, i-1
			shuffle aLetters, 0, i-1
		END IF
		
		SET oFolder = Nothing
	END IF

END SUB



SUB buildTextListIndex( sFolder, sIndexName )

	DIM	aJokes()
	REDIM aJokes(50)
	DIM	sJokesStream
	
	readTextListFolder sFolder, aJokes
	
	sJokesStream = JOIN( aJokes, vbCRLF )

	IF g_oFSO.FileExists( sIndexName ) THEN
		ON Error Resume Next
		g_oFSO.DeleteFile sIndexName, TRUE
		ON Error GOTO 0
	END IF
	
	ON Error Resume Next
	DIM	oFile
	SET oFile = g_oFSO.CreateTextFile( sIndexName, TRUE )
	IF NOT Nothing IS oFile THEN
		oFile.Write sJokesStream
		oFile.Close
		SET oFile = Nothing
	END IF
	ON Error GOTO 0

END SUB



FUNCTION needToRebuildIndex( sBaseName )

	needToRebuildIndex = FALSE

	DIM	dIndexTime
	DIM	nIndexCount
	DIM	nExpire
	
	dIndexTime = CDATE(EVAL( "g_d" & sBaseName & "IndexTime" ))
	nIndexCount = CINT(EVAL( "g_n" & sBaseName & "IndexCount" ))
	nExpire = nIndexCount / 2.5
	IF 7*8 < nExpire THEN nExpire = 7*8
	IF nExpire < ABS(DATEDIFF( "d", dIndexTime, NOW )) THEN
		DIM	sFolder
		DIM	sIndex
		sFolder = EVAL( "g_s" & sBaseName & "Folder" )
		sIndex = cache_filepath( "shuffle", sBaseName&".txt" )
		'sIndex = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
		
		Response.Write "<span style=""color:gray"">" & LCASE(LEFT(sBaseName,1)) & "</span>"
		IF Response.Buffer THEN Response.Flush
		buildTextListIndex sFolder, sIndex
		needToRebuildIndex = TRUE
	END IF

END FUNCTION




FUNCTION getTextListFilename( sBaseName )

	getTextListFileName = ""

	DIM	sFile
	sFile = findShuffleIndex( sBaseName )
	IF "" <> sFile THEN
		DIM	oIndexFile
		SET oIndexFile = g_oFSO.GetFile( sFile )
		IF NOT Nothing IS oIndexFile THEN
			EXECUTEGLOBAL "g_d" & sBaseName & "IndexTime = CDATE(""" & oIndexFile.DateLastModified & """)"
			
			DIM oStream
			SET oStream = oIndexFile.OpenAsTextStream( 1 )	' read only
			IF NOT Nothing IS oStream THEN
				DIM	xData
				DIM	aData
				xData = oStream.ReadAll
				oStream.Close
				SET oStream = Nothing
				aData = SPLIT( xData, vbCRLF )
				
				EXECUTEGLOBAL "g_n" & sBaseName & "IndexCount = " & CSTR(UBOUND(aData) - LBOUND(aData) + 1)
				
				DIM	i
				i = getNextCookie( sBaseName, LBOUND(aData), UBOUND(aData) )
				IF -1 < i THEN
					DIM	sPath
					DIM	sFolder
					sFolder = EVAL( "g_s" & sBaseName & "Folder" )
					sFile = aData(i)
					sPath = g_oFSO.BuildPath( sFolder, sFile )
					IF g_oFSO.FileExists( sPath ) THEN
						getTextListFilename = sPath
					END IF
				END IF
			END IF
			SET oIndexFile = Nothing
		END IF
	END IF

END FUNCTION







FUNCTION needToRebuildGalleryIndex()

	needToRebuildGalleryIndex = FALSE

	DIM	dIndexTime
	DIM	nIndexCount
	DIM	nExpire
	
	dIndexTime = CDATE(g_dGalleryIndexTime)
	nIndexCount = CINT(g_nGalleryIndexCount)
	nExpire = nIndexCount / (3*5)
	'Response.Write "Gallery Expire = " & nExpire & "<br>"
	IF 7*8 < nExpire THEN nExpire = 7*8
	IF nExpire < ABS(DATEDIFF( "d", dIndexTime, NOW )) THEN
		DIM	sFolder
		DIM	sIndex
		sFolder = g_sGalleryFolder
		sIndex = cache_filepath( "shuffle", "gallery.txt" )
		'sIndex = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
		
		Response.Write "<span style=""color:gray"">g</span>"
		IF Response.Buffer THEN Response.Flush
		buildGalleryIndex sFolder, sIndex
		needToRebuildGalleryIndex = TRUE
	END IF

END FUNCTION



FUNCTION hasGalleryFilenames()
	DIM	i
	hasGalleryFilenames = TRUE
	FOR i = LBOUND(g_aGalleryFilenames) TO UBOUND(g_aGalleryFilenames)
		IF "" = g_aGalleryFilenames(i) THEN
			hasGalleryFilenames = FALSE
		END IF
	NEXT
END FUNCTION



SUB getGalleryFilenames()

	DIM	sBaseName
	sBaseName = "Gallery"
	
	DIM	sFile
	sFile = findShuffleIndex( sBaseName )
	IF "" <> sFile THEN
		DIM	oIndexFile
		SET oIndexFile = g_oFSO.GetFile( sFile )
		IF NOT Nothing IS oIndexFile THEN
			g_dGalleryIndexTime = oIndexFile.DateLastModified
			
			DIM oStream
			SET oStream = oIndexFile.OpenAsTextStream( 1 )	' read only
			IF NOT Nothing IS oStream THEN
				DIM	xData
				DIM	aData
				xData = oStream.ReadAll
				oStream.Close
				SET oStream = Nothing
				aData = SPLIT( xData, vbCRLF )
				
				IF  1024 < LEN(xData)  AND  LBOUND(aData) < UBOUND(aData) THEN
					g_nGalleryIndexCount = UBOUND(aData) - LBOUND(aData) + 1
					
					DIM	i
					DIM	j
					DIM	n
					DIM	sLine
					DIM	aLine
					DIM	sSuffix
					j = 0
					DO WHILE j <= UBOUND(g_aGalleryFilenames)
						i = getNextCookie( "Gallery", LBOUND(aData), UBOUND(aData) )
						IF -1 < i THEN
							sLine = aData(i)
							IF 0 < LEN(sLine) THEN
								aLine = SPLIT( sLine, vbTAB )
								n = InStrRev( aLine(1), ".", -1, vbTextCompare )
								IF 0 < n THEN
									sSuffix = LCASE(MID( aLine(1), n ))
									SELECT CASE sSuffix
									CASE ".gif", ".jpg", ".jpeg", ".png"
										g_aGalleryFilenames(j) = aLine(2)
										j = j + 1
									END SELECT
								END IF
							END IF
						END IF
					LOOP
				END IF
			END IF
			SET oIndexFile = Nothing
		END IF
	END IF

END SUB


FUNCTION getHumorFilename()
	getHumorFilename = getTextListFilename( "Humor" )
END FUNCTION

FUNCTION getSafetyFilename()
	getSafetyFilename = getTextListFilename( "Safety" )
END FUNCTION

g_sHumorFilename = getHumorFilename()
g_sSafetyFilename = getSafetyFilename()
'getGalleryFilenames














%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<!--DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & Server.HTMLEncode(g_sSiteName & " - " & g_sShortSiteName & " - " & g_sSiteCity & " - GWRRA" ) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Home - Default</title>
<%
END IF
%>
<meta name="keywords" content="gwrra, goldwing, gold wing road riders association, <%=g_sSiteName%>, honda, motorcycle, <%=g_sSiteCity%>, <%=g_sShortSiteName%>">
<meta name="navigate" content="tab,home">
<meta name="navtitle" content="Home">
<meta name="sortname" content="aaaaaaaaaaaaaaaaadefault">
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<script type="text/javascript" language="javascript" src="scripts/slideshow.js"></script>
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







<%

IF hasGalleryFilenames() THEN


SUB makeJSImageList( sName, aList() )

	DIM	nList
	nList = UBOUND(aList) + 1

	Response.Write "var g_n" & sName & " = " & nList & ";" & vbCRLF
	Response.Write "var g_a" & sName & " = new Array(" & nList & ");" & vbCRLF

	DIM xFile
	DIM	i
	i = 0
	FOR EACH xFile IN aList
		Response.Write "g_a" & sName & "[" & i & "] = """ & xFile & """;" & vbCRLF
		i = i + 1
	NEXT 'xFile

END SUB

makeJSImageList "Gallery", g_aGalleryFilenames


%>




CSlidePicture.prototype.precook
= function()
{
	var	bResult = true;
	if ( ! this.m_bCached )
	{
		this.m_oImage.src = 'picture_old.asp?file='+this.m_source;
		this.m_bCached = true;
	}
	if ( ! this.m_oImage.complete )
		bResult = false;
	return bResult;
};





CSlidePicture.prototype.projectSlide
= function( nHeight, nWidth )
{
	var h = this.height();
	var w = this.width();
	
	if ( 0 < h  &&  0 < w )
	{
		var WidthRatio  = w / nWidth;
		var HeightRatio = h / nHeight;
		var Ratio = WidthRatio > HeightRatio ? WidthRatio : HeightRatio;
		
	
		w  = Math.floor(w / Ratio);
		h = Math.floor(h / Ratio);
	}
	else
	{
		h = nHeight;
		w = nWidth;
	}
	
	var s = '<div class="SlideShowImg">'
		+	'<a href="gallery_findpicture.asp?file='+this.m_source+'">'
		+	'<img src="picture_old.asp?file='+this.m_source+'" border="0" height="'+h+'" width="'+w+'" >'
		+	'</a>'
		+	'</div>';

	return s;
};



CSlideScreenTransitionRandom.prototype.watchObjectSize
= function( o )
{
	var tempThis = this;
	o.onresize = (function(e){e = e || window.event;tempThis["handleOnResizeEnd"](e,this)});
}


CSlideScreenTransitionRandom.prototype.handleOnResizeEnd
= function( e )
{
	var o = e.srcElement;
	var oBox = new CSlideBox()
	oBox.setBox( o );
	var h = oBox.height();
	var w = oBox.width();
	var nHBox = Math.floor(w * 5 / 7);
	if ( h != nHBox )
	{
		oBox.setHeight( nHBox );
	}
	if ( w != this.width() )
		this.setWidth( w );
	if ( nHBox != this.height() )
		this.setHeight( nHBox );
	return true;
}


function handleResizeForSlideScreen( e )
{
	var o;
	var oBox;
	var w;
	var	b = false;
	o = document.getElementById("colleft");
	oBox = new CSlideBox()
	oBox.setBox( o );
	w = oBox.width();
	if ( w < 280 )
	{
		oBox.setWidth( 280 );
		b = true;
	}
	else if ( 300 < w )
	{
		oBox.setWidth( 300 );
		b = true;
	}
	
	o = document.getElementById( "colright" );
	oBox.setBox( o );
	w = oBox.width();
	if ( w < 190 )
	{
		oBox.setWidth( 190 );
		b = true;
	}
	else if ( 200 < w )
	{
		oBox.setWidth( 200 );
		b = true;
	}

	o = document.getElementById( "colmiddle" );
	oBox.setBox( o );
	w = oBox.width();
	if ( w < 220 )
	{
		oBox.setWidth( 220 );
		b = true;
	}
	
	if ( b )
		return true;

	
	o = document.getElementById( "slideframe" );
	oBox = new CSlideBox()
	oBox.setBox( o );
	var h = oBox.height();
	w = oBox.width();
	var nHBox = Math.floor(w * 5 / 7);
	if ( h != nHBox )
	{
		oBox.setHeight( nHBox );
	}
	if ( w != g_oSlideScreen.width() )
		g_oSlideScreen.setWidth( w );
	if ( nHBox != g_oSlideScreen.height() )
		g_oSlideScreen.setHeight( nHBox );
	return true;
}





var g_oCarousel = new CSlideCarousel();
g_oCarousel.load( g_aGallery );


var g_oProjector = new CSlideProjectorShowSingle();
g_oProjector.setCarousel( g_oCarousel );
g_oProjector.setPicturePause( <%=g_nPictureDelay%> );


var g_oSlideScreen;







function startSlideShow()
{

	var oSlideFrame = document.getElementById( "slideframe" );
	var oBox = new CSlideBox();
	oBox.setBox( oSlideFrame );
	
	g_oSlideScreen = new CSlideScreenTransitionRandom( "slidescreen" );
	g_oSlideScreen.initialize();
	
	//g_oSlideScreen.watchObjectSize( oSlideFrame );
	
	var w = oBox.width();
	var h = Math.floor(w * 5 / 7);
	
	oBox.setHeight( h );
	g_oSlideScreen.setWidth(w);
	g_oSlideScreen.setHeight(h);


	if (window.addEventListener)
		window.addEventListener("resize", handleResizeForSlideScreen, false);
	else if (window.attachEvent)
		window.attachEvent("onresize", handleResizeForSlideScreen);

	
	
	//setupSlideScreen();
	g_oProjector.setScreen( g_oSlideScreen );
	

	g_oCarousel.precookCards();
	//currentslide = 0;
	setTimeout( "g_oProjector.run()", 0.01*1000 );
}



if (window.addEventListener)
	window.addEventListener("load", startSlideShow, false);
else if (window.attachEvent)
	window.attachEvent("onload", startSlideShow);
else
	setTimeout( "startSlideShow()", 2.0*1000 );


<%
END IF

%>





//-->
</script>
<style type="text/css">
<!--
.rssweather
{
	text-align: left;
	font-size: small;
	font-family: sans-serif;
}

.meetcalendar
{
	display: none;
}

#slideframetable
{
	position: relative;
	width: 100%;
}




#slideframe
{
	position: relative;
	width: 100%;
	padding: 0;
	margin: 0;
	border: 0;
	background-image: url('themes/default/bkg_gallery.jpg');
	overflow: hidden;
}


#slidescreen
{
	position: relative;
	height: 100px;
	width: 100%;
	filter: revealTrans(Duration=1,Transition=23);
	overflow: hidden;
}


#slidescreen .tempPageLoading
{
	color: white;
	text-align: center;
}


#slidescreen div.SlideShowImg
{
	position: relative;
	width: 100%;
	text-align: center;
}


#slidescreen img
{
	border: 0;
}


.neighbormeeting
{
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: x-small;
}







-->
</style>
<%
makeWeatherStyles
%>
</head>

<body onload="doFramesBuster()" class="homepage">

<!--#include file="scripts\page_begin.asp"-->
<!--#include file="scripts\include_navigation.asp"-->
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
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="tableoutercolset">
<tr>
<td id="colleft">
<%
'========================
' ColLeft
'========================
IF FALSE THEN
%>
	<p>ColLeft</p>
<%
END IF






RSSForum "forum.xml", g_sDomain, "icon_forum.gif", g_sSiteTitle & " Forum"



%>
<!--div align="center"-->
<div class="announcements">
<%

	DIM	dAnnouncementsFile
	'DIM	dAnnouncementsModified
	DIM	dAMTemp
	DIM	aAnnouncements()
	DIM	nAnnouncements
	
	'g_sUseFileNameSuffix = ".htm"
	nFileCount = 0
	'buildFileList "announcements", TRUE
	appendFileList 0, -21
	
	nAnnouncements = nFileCount
	REDIM aAnnouncements(UBOUND(aFileList))
	FOR nLen = 0 TO UBOUND(aFileList)
		aAnnouncements(nLen) = aFileList(nLen)
	NEXT 'nLen
	
	IF 0 < nFileCount THEN
	
		aFileSplit = SPLIT( aFileList(0), vbTAB )
		dAnnouncementsModified = CDATE( aFileSplit(kFI_DateLastModified) )
	

		DIM	sJPGFile
		DIM	sSrcUpdate
		DIM	sAltUpdate
		FOR nLen = 0 TO nFileCount-1
			aFileSplit = SPLIT( aFileList(nLen), vbTAB, -1, vbTextCompare )
			dAnnouncementsFile = CDATE(aFileSplit(kFI_DateLastModified))
			IF 0 < DATEDIFF( "s", g_dLastVisit, dAnnouncementsFile) THEN
				sSrcUpdate = "Update"
				sAltUpdate = "Updated Announcement/Extra"
			ELSE
				sSrcUpdate = ""
				sAltUpdate = "Announcement/Extra"
			END IF
%>
<!--table border="1" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" 
width="268" style="border-collapse: collapse" bordercolor="#FFCC00">
<tr>
<th bgcolor="#FFDD99"-->
<div class="announcementstitle">
<img border="0" src="<%=g_sTheme%>icon_announcement<%=sSrcUpdate%>.gif" alt="<%=sAltUpdate%>" align="absbottom"><%
			dAMTemp = CDATE( aFileSplit(kFI_DateLastModified) )
			IF dAnnouncementsModified < dAMTemp THEN dAnnouncementsModified = dAMTemp
			'Response.Write "<a href=""announcements.asp?page=" & aFileSplit(kFI_Name) & """>"
			Response.Write Server.HTMLEncode(aFileSplit(kFI_Title)) 
			'Response.Write "</a>"
%>
</div>
<!--/th>
</tr>
<tr>
<td-->
<%
			sJPGFile = ""
			sJPGFile = REPLACE( aFileSplit(kFI_Name), ".htm", ".jpg" )
			IF NOT g_oFSO.FileExists( Server.MapPath( "announcements/" & sJPGFile ) ) THEN
				sJPGFile = REPLACE( aFileSplit(kFI_Name), ".htm", ".gif" )
				IF NOT g_oFSO.FileExists( Server.MapPath( "announcements/" & sJPGFile ) ) THEN
					sJPGFile = REPLACE( aFileSplit(kFI_Name), ".htm", ".png" )
					IF NOT g_oFSO.FileExists( Server.MapPath( "announcements/" & sJPGFile ) ) THEN
						sJPGFile = ""
					END IF
				END IF
			END IF
			IF "" <> sJPGFile THEN
				Response.Write "<div align=""center"">"
				'Response.Write "<a href=""announcements.asp?page=" & aFileSplit(kFI_Name) & """>"
				Response.Write "<img border=""0"" alt=""" & Server.HTMLEncode(aFileSplit(kFI_Title)) & """ src=""announcements/" & sJPGFile & """>" & vbCRLF
				'Response.Write "</a>"
				Response.Write "</div>" & vbCRLF
			END IF
			IF 0 < LEN(aFileSplit(4)) THEN
				'Response.Write "<br>"
				Response.Write "<div align=""left"">"
				Response.Write Server.HTMLEncode(aFileSplit(4))
				'Response.Write "&nbsp;&nbsp; <a href=""announcements.asp?page=" & aFileSplit(kFI_Name) & """>"
				'Response.Write "More..."
				'Response.Write "</a>"
				Response.Write "</div>" & vbCRLF
			END IF
			'Response.Write "<div style=""text-align: left; color: #999999; font-family: sans-serif; font-size: xx-small;"">"
			'Response.Write "Updated: " & Server.HTMLEncode(DATEADD("h", g_nServerTimeZoneOffset, dAnnouncementsFile))
			'Response.Write "</div><br>"
%>
<!--/td>
</tr-->
<%
		NEXT 'nLen
	END IF









makeAnnouncements
%>
</div>
<!--/div-->
<%
RSSForum "districtforum.xml", g_sDistrictDomain, "icon_alforum.gif", "District Forum"

%>
<!--div align="center"-->
<div class="announcements">
<%


RSSDistrictAnnouncements

%> 
</div>
<!--/div-->


</td>
<td id="colmiddle">
<%
'========================
' ColMiddle
'========================
IF FALSE THEN
%>
	<p>ColMiddle</p>
<%
END IF


IF hasGalleryFilenames() THEN

%>
<table border="0" cellpadding="4" cellspacing="0" width="100%">
	<tr>
		<th bgcolor="#FFDD99" class="galleryheader" style="border-top: 1px solid #FFFFFF" bgcolor="#FFDD99" align="left" background="<%=g_sTheme%>header_bkg.jpg">
			<img border="0" src="<%=g_sTheme%>Icon_Camera_s.gif" alt="Gallery" align="absbottom"> 
			From the <a href="gallery.asp">Gallery</a>&nbsp;&nbsp; <span style="font-size: smaller;">
			of <%=g_nGalleryIndexCount%> pictures</span></th>
	</tr>
</table>
<div id="slideframe">
<div id="slidescreen">
<a href="gallery_findpicture.asp?file=<%=Server.URLEncode( g_aGalleryFilenames(UBOUND(g_aGalleryFilenames)) )%>"><img width="100%" border="0" alt="Gallery Pictures" src="picture_old.asp?file=<%=Server.URLEncode( g_aGalleryFilenames(UBOUND(g_aGalleryFilenames)) )%>"></a>
</div>
</div>

<%
END IF
%>
<%

DIM	dNow
dNow = DATEADD( "h", g_nServerTimeZoneOffset, NOW)


DIM	nDateBegin
DIM	nDateEnd


DIM	sRmdHtm


nDateBegin = jdateFromVBDate( dNow )
nDateEnd = nDateBegin + 7 * 12


sRmdHtm = remind_cache( "homegathering.htm", "h", 24, "d", _
					nDateBegin, nDateEnd, LCASE(g_sSiteChapterID) & "-gathering", "ignore", FALSE, dNow, _
					Server.MapPath("scripts/remind_gathering.xslt") )

IF "" <> sRmdHtm THEN


%>
		
		<table width="100%" cellpadding="4" cellspacing="0" border="0">
			<tr>
				<th bgcolor="#FFDD99" align="left" background="<%=g_sTheme%>header_bkg.jpg">
				<img border="0" src="<%=g_sTheme%>icon_standardevents.gif" alt="Monthly Gathering" align="absbottom"> 
				Next Monthly Gathering</th>
			</tr>
			<tr>
				<td>
<%

Response.Write sRmdHtm



%>
				</td>
			</tr>
		</table>
<%

END IF

'===========================================================




nDateBegin = jdateFromVBDate( dNow )
nDateEnd = nDateBegin + 7 * 3 - 1

sRmdHtm = remind_cache( "homeevents.htm", "h", 24, "d", _
					nDateBegin, nDateEnd, LCASE(g_sSiteChapter) & ",core,key", "holiday,usa,none", TRUE, dNow, _
					Server.MapPath("scripts/remind.xslt") )

IF "" <> sRmdHtm THEN


%>
		<table border="0" cellpadding="4" cellspacing="0" width="100%">
			<tr>
				<th bgcolor="#FFDD99" align="left" background="<%=g_sTheme%>header_bkg.jpg">
				<img border="0" src="<%=g_sTheme%>icon_events.gif" alt="Events" align="absbottom">
				<a href="calendar.asp?cat=<%=LCASE(g_sSiteChapter)%>">Chapter Events / Rides</a></th>
			</tr>
		</table>
<%

	Response.Write sRmdHtm


END IF
'===========================================================





DIM	nYear, nYearEnd
DIM	nMon, nMonEnd


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

sRmdHtm = remind_cache( "homebday.htm", "d", 7, "m", _
					nDateBegin, nDateEnd, sRmdCat, "ignore", FALSE, dNow, _
					Server.MapPath("scripts/remind.xslt") )

IF "" <> sRmdHtm THEN

%>
		<table border="0" cellpadding="4" cellspacing="0" width="100%">
			<tr>
				<th bgcolor="#FFDD99" align="left" background="<%=g_sTheme%>header_bkg.jpg">
				<img border="0" src="<%=g_sTheme%>icon_events.gif" alt="Birthdays" align="absbottom">
				<a href="calendar.asp?cat=<%=sRmdCat%>">Birthdays / Anniversaries</a></th>
			</tr>
		</table>
<%

	Response.Write sRmdHtm


END IF

IF ISDATE( dRemindLastModified ) THEN

%>
<p style="color:#999999; font-family: sans-serif; font-size: xx-small;">Calendar 
Updated: <%=DATEADD("h", g_nServerTimeZoneOffset, dRemindLastModified)%></p>
<%

END IF


'===========================================================



%>
		<table border="0" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" width="100%">
			<tr>
				<th bgcolor="#FFDD99" align="left" background="<%=g_sTheme%>header_bkg.jpg">
				<img border="0" src="<%=g_sTheme%>icon_email.gif" alt="Email Events" align="absbottom">
				<a href="notification_upcoming.asp">Chapter Event Notification</a></th>
			</tr>
			<tr>
				<td><a href="notification_upcoming.asp">Sign-up</a> for email notification of upcoming 
				Chapter Rides &amp; Events. </td>
			</tr>
		</table>

</td>
<td id="colright">
<%
'========================
' ColRight
'========================
IF FALSE THEN
%>
	<p>ColRight</p>
<% 
END IF


pagebody_saveFuncs
outputMultiplePages TRUE	' TRUE to generate an enclosing table
pagebody_restoreFuncs


IF "" <> g_sSiteZip THEN
%>
		<table width="100%" cellspacing="0" cellpadding="4">
			<tr>
				<th bgcolor="#99CCFF" align="left" background="<%=g_sTheme%>header_bkg_weather.jpg">
				<img border="0" src="<%=g_sTheme%>icon_weather.gif" alt="Weather" align="absbottom">
				<a href="http://grizzlyweb.com/links/weather_forecast.asp" target="_blank">
				Weather</a></th>
			</tr>
		</table>
		<%



rssweather g_sSiteZip

END IF
%>
<table border="0" cellpadding="4" cellspacing="0" bgcolor="#EEEEEE">
<%

IF Response.Buffer THEN Response.Flush



FUNCTION jokePicture( sLabel )

	DIM	sFolder
	sFolder = g_sHumorFolder
	jokePicture = virtualFromPhysicalPath(g_oFSO.BuildPath(sFolder, "images\" & sLabel))
	jokePicture = "picture_old.asp?file=" & Server.URLEncode(jokePicture)

END FUNCTION


FUNCTION processJoke()

	processJoke = ""

	DIM	oFile
	DIM	oF
	
	IF "" <> g_sHumorFilename THEN
	
		SET oF = g_oFSO.GetFile( g_sHumorFilename )
		
		SET oFile = oF.OpenAsTextStream( 1 )	'open for read
		
		DIM	sText
		sText = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
		SET oF = Nothing
	
		DIM	sSavePictureFunc
		sSavePictureFunc = g_htmlFormat_pictureFunc
		g_htmlFormat_pictureFunc = "jokePicture"
		
		gHtmlOption_encodeEmailAddresses = TRUE
		processJoke = htmlFormatCRLF( sText )
		
		g_htmlFormat_pictureFunc = sSavePictureFunc
	
	END IF
	
END FUNCTION

DIM	sJoke
sJoke = processJoke()
IF "" <> sJoke THEN


%>
<tr>
<th bgcolor="#FFDD99" align="left" background="<%=g_sTheme%>header_bkg.jpg">
	<img border="0" src="<%=g_sTheme%>icon_smile.gif" alt="Smile" align="absbottom">Smile</th>
</tr>
<tr>
<td bgcolor="#FFEECC" class="SmileBody">
<%
	Response.Write sJoke
%>
</td>
</tr>
<%
END IF


FUNCTION safetyPicture( sLabel )

	DIM	sFolder
	sFolder = g_sSafetyFolder
	safetyPicture = virtualFromPhysicalPath(g_oFSO.BuildPath(sFolder, "images\" & sLabel))
	safetyPicture = "picture_old.asp?file=" & Server.URLEncode(safetyPicture)

END FUNCTION



FUNCTION processSafetyHints()

	processSafetyHints = ""

	DIM	oFile
	DIM	oF
	
	IF "" <> g_sSafetyFilename THEN
	
		SET oF = g_oFSO.GetFile( g_sSafetyFilename )
		
		SET oFile = oF.OpenAsTextStream( 1 )	'open for read
		
		DIM	sText
		sText = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
		SET oF = Nothing
	
		DIM	sSavePictureFunc
		sSavePictureFunc = g_htmlFormat_pictureFunc
		g_htmlFormat_pictureFunc = "safetyPicture"
		
		gHtmlOption_encodeEmailAddresses = TRUE
		processSafetyHints = htmlFormatCRLF( sText )
		
		g_htmlFormat_pictureFunc = sSavePictureFunc
	
	END IF

END FUNCTION

SUB genSafetyHint()

DIM	sHints
sHints = processSafetyHints()

IF "" <> sHints THEN

%>
<!--table border="0" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" class="SafetyTips" style="border-collapse: collapse"-->
	<tr>
		<th bgcolor="#99FF99" align="left" class="SafetyTips">
			<img border="0" src="<%=g_sTheme%>icon_safety.gif" alt="Safety" align="absbottom">Riding 
			Safety Tips</th>
	</tr>
	<tr>
<td align="left" bgcolor="#CCFFCC" class="SafetyTips">
<%
	Response.Write sHints
%>
</td>
	</tr>
<!--/table-->

<%
END IF

END SUB
genSafetyHint




'===========================================================




nDateBegin = jdateFromVBDate( dNow )
nDateEnd = nDateBegin + 7 * 8

sRmdHtm = remind_cache( "homeneighbors.htm", "h", 24, "d", _
					nDateBegin, nDateEnd, "neighbor-meeting", "ignore", FALSE, dNow, _
					Server.MapPath("scripts/remind_neighbors.xslt") )

IF "" <> sRmdHtm THEN


%>
<tr>
	<th bgcolor="#FFDD99" background="images/header_bkg.jpg" align="left"><img border="0" src="<%=g_sTheme%>icon_gwrra.gif" alt="GWRRA" align="absbottom"> Neighbors</th>
</tr>
<tr>
<td>
<%

	Response.Write sRmdHtm

%>
<p>Check the 
		<a target="_blank" href="http://www.alabama-gwrra.org/">AL</a> and 
		<a target="_blank" href="http://www.tngwrra.org/">TN</a> GWRRA Web Pages 
		for other chapters of interest</p>

</td>
</tr>
<%
END IF
'===========================================================
%>
</table>


</td>
</tr>
</table>

<hr>
<!--#include file="scripts\page_end.asp"--><%
IF g_dLocalPageDateModified < dRemindLastModified THEN g_dLocalPageDateModified = dRemindLastModified
%>
<p style="font-size: x-small; color: silver">Last Visit: <%=DATEADD("h", g_nServerTimeZoneOffset, g_dLastVisit)%>, 
Active Visitors: <%=Application("ActiveUsers")%></p>
<%


IF NOT needToRebuildIndex( "Safety" ) THEN
	IF NOT needToRebuildIndex( "Humor" ) THEN
		IF NOT needToRebuildGalleryIndex() THEN
		END IF
	END IF
END IF
Response.Write "<br>"


IF Response.Buffer THEN Response.Flush
%>
<!--#--in--clude file="scripts\include_emailnotification.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
