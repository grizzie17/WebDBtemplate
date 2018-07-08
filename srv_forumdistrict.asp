<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts\rss.asp"-->
<!--#include file="scripts\include_forum.asp"-->
<%

RSSForum "districtforum.xml", g_sDistrictDomain, "icon_alforum.gif", "District Forum"



SET g_oFSO = Nothing
%>
