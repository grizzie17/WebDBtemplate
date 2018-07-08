<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Clear Cache</title>
<script type="text/javascript" language="JavaScript">
<!--

var sTargetURL = "./";

function doRedirect()
{
    setTimeout( "timedRedirect()", 2*1000 );
}

//  There are two definitions of 'timedRedirect', this
//  one adds to the visitors page history.
function timedRedirect()
{
    window.location.href = sTargetURL;
}

//-->
</script>

<script type="text/javascript" language="JavaScript1.1">
<!--

function timedRedirect()
{
    window.location.replace( sTargetURL );
}

//-->
</script>

</head>

<body>

<%

cache_clearFolders

%>

<p>Cache Cleared</p>

<script type="text/javascript" language="javascript">

doRedirect();

</script>

</body>

</html>
<%
SET g_oFSO = Nothing
%>