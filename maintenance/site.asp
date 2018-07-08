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
<!--#include file="scripts\htmlformat.asp"-->
<%










%><html>

<head>
<meta name="navigate" content="tab">
<title>Site</title>
<meta name="sortname" content="zzzzzzzzzzzsite" />
<link rel="stylesheet" href="theme/style.css" type="text/css">


</head>

<body bgcolor="#FFFFFF">
<!--#include file="scripts/index_include.asp"-->

<div style="text-align: right">Editors for maintaining Site configuration and layout</div>

<%
IF authHasAccess("CFGe")  OR  authHasAccess("SRVe") THEN
%>

<ul>
	<li><a href="sitelayoutedit.asp?type=screen">Edit Homepage Layout</a></li>
	<li><a href="sitelayoutedit.asp?type=mobile">Edit Mobile Homepage Layout</a></li>
	</ul>
	<ul>
<li><a href="configedit.asp">Edit Site Configuration</a></li>
</ul>

<%	
END IF

IF authHasAccess("USRm") THEN
%>
<ul>
<li><a href="users.asp">Maintain Users</a></li>
</ul>
<%
END IF
IF authHasAccess("SRVe") THEN
%>
<p style="margin-left: 2em;"><span style="font-size: larger; font-weight: bold">WARNING</span><br />
Do not select the following 'DB' links as they may erase the database and reset it to its original state.<br />
Make sure you know what you are doing.</p>
<ul>
	<li><a target="_blank" href="dbupgrade.asp">DB Upgrade</a> (opens new window)</li>
	<li><a target="_blank" href="dbinfo.asp">DB Info</a> (opens new window)</li>
</ul>

<%
END IF


%>
<p><a href="default.asp">Back to Maintenance Page</a></p>
<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->