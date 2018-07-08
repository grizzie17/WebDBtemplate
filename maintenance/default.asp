<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Server.ScriptTimeout = 60*30
Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<html>

<head>
<meta name="navigate" content="tab">
<meta name="navtitle" content="Home">
<title>Content Maintenance</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="sortname" content="aaaadefault.asp">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<meta name="robots" content="noindex">
</head>
<body>
<!--#include file="scripts/index_include.asp"-->



<ul>
<%
IF authHasAccess( "CALm" ) THEN
%>
	<li><a target="_blank" href="remind.asp">Calendar</a>&nbsp; (opens new window)</li>
<%
END IF
%>
	<li><a href="helpCalendar.htm" target="_blank">Calendar Help</a></li>
    <li><a href="clearcache.asp">Clear Cache</a> (needed for many calendar edits 
	to show up on home-page)</li>
	<li><a href="cleanremind.asp" target="_blank">Clean Calendar Scripts</a> 
	(used to remove attack scripts - opens new window)</li>
</ul>

<ul>
<li><a target="_blank" href="newsletternotice.asp">Newsletter Notice</a> (opens 
new window)</li>
<li><a target="_blank" href="../webalizer">Web Statistics</a> (opens new window)</li>
</ul>


<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->