<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts/include_db_begin.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts/include_navtabs.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<%


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & g_sNavPageTitle & " - " & Server.HTMLEncode(g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Gallery</title>
<%
END IF
%>
<link rel="shortcut icon" href="./favicon.ico">
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
</head>

<body>
<!--#include file="scripts\page_begin.asp"-->
<div class="noprint">
<!--#include file="app_data\layout\header.asp"-->
</div>

<!--#include file="scripts\page_end.asp"-->
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

<!--webbot bot="Include" i-checksum="20551" endspan -->
</body>

</html>
<!--#include file="scripts/include_db_end.asp"-->
