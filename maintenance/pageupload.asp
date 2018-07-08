<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Server.ScriptTimeout = 60*30

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%
picture_createObject




picture_close

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "pages.asp?category=" & sQualCategory

%>