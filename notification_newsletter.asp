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
<title>Newsletter Notification - <%=Server.HTMLEncode(g_sSiteName)%></title>
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<style type="text/css">

iframe
{
	background-image: url(images/bkg_sample.jpg);
}




</style>
</head>

<body>

<!--#include file="scripts\page_begin.asp"-->
<!--#include file="app_data\layout\header.asp"-->

<p>Sign-up for email notification of Newsletters.</p>
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
			<input type="hidden" name="list" value="NewsletterNotice">
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