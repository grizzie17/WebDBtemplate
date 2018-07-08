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
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & g_sNavPageTitle & " - " & Server.HTMLEncode(g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Forum</title>
<%
END IF
%>
<!--#include file="scripts\favicon.asp"-->
<link rel="stylesheet" title="Default Styles" type="text/css" href="<%=g_sTheme%>style.css">
<meta name="navigate" content="tab">
<meta name="navtitle" content="Forum">
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
</head>

<body>


<!--#include file="scripts\page_begin.asp"-->
<!--#include file="config\header.asp"-->

<p>The <a target="_blank" href="http://Forum.<%=g_sDomain%>"><%=Server.HTMLEncode(g_sSiteName)%> Forum</a> 
opens in a separate window.&nbsp; You are required to register to post messages.</p>
<p>If you are having trouble getting registered on the <%=Server.HTMLEncode(g_sSiteName)%> 
forum, please register; then please send an email to the
<span class="pobox">webmaster[web ma ster|<%=g_sDomain%> - Forum Registration]</span> 
containing the following: requested &quot;username&quot;, registered &quot;email address&quot;, and your 
real name.&nbsp; I apologize for asking for 
so much information but we receive a significant number of requested 
registrations weekly.&nbsp; Usually all of them are spammers.</p>

<p>&gt;&gt;&gt; <b> <font size="4"> <a target="_blank" href="http://Forum.<%=g_sDomain%>"><%=Server.HTMLEncode(g_sSiteName)%> 
Forum</a> </font></b></p>
<p>&nbsp;</p>

<p><font color="#808080">No registration information is EVER provided to mass 
mailing schemes.</font></p>

<br clear="all">
<!--#include file="scripts\page_end.asp"-->


</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
