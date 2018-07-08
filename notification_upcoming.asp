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
<!--#include file="scripts\include_theme.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<link rel="shortcut icon" href="favicon.ico">
<link rel="stylesheet" title="Default Styles" type="text/css" href="<%=g_sTheme%>style.css">
<title>Calendar Event Notification - <%=Server.HTMLEncode(g_sSiteName)%></title>
<meta name="navigate" content="upcoming-events">
<meta name="navtitle" content="Calendar Event Notification">
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<style type="text/css">

iframe
{
	background-image: url(images/bkg_sample.jpg);
}




</style>
</head>

<body>

<%
DIM g_sNavigateTabLabel
g_sNavigateTabLabel = "(home|upcoming-events)"

%>
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="app_data\layout\header.asp"-->

<p>Sign-up for email notification of upcoming Ride/Schedule events.&nbsp; The notification also 
includes announcements/articles that have 
recently been added or changed.&nbsp; You will receive an email similar to the 
sample provided below.</p>
<p>To register or remove your registration simply use the following form.&nbsp; A confirmation 
email will be sent that you must reply to before the requested action will take effect.</p>
		<form method="POST" action="notification.asp">
			<table>
				<tr>
					<td>Email:</td>
					<td><input type="text" name="email" size="30"></td>
				</tr>
				<tr>
					<td>Action:</td>
					<td>
			<input type="hidden" name="list" value="UpcomingEvents">
			<select name="action">
			<option value="register">Register</option>
			<option value="unregister">Unregister</option>
			</select></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td>
			<input type="submit" value="Submit" name="B1"></td>
				</tr>
			</table>
		</form>

<iframe name="SampleEventNotify" src="sampleeventnotify.asp" marginheight="5" width="95%" title="Sample Event Notification" height="500" target="_top" style="border: 2px solid #000000">Your browser does not support inline frames or is currently configured not to display inline frames.</iframe>
<p>&nbsp;</p>

<p><font color="#808080">No registration information is EVER provided to mass 
mailing schemes.</font></p>

<br clear="all">
<!--#include file="scripts\page_end.asp"-->


</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
<%
%>