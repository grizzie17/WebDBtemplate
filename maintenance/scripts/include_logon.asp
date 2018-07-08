<!--#include file="include_logon_defs.asp"-->
<%
'depends upon config.asp

'include file that assures user is logged on


g_auth_sUsername = authCookie( "Username" )
g_auth_sPassword = authCookie( "PW" )
g_auth_nUserID = authCookieNumber( "UserID" )
g_auth_bAccept = authCookieBool( "Accept" )
g_auth_sFlags = authCookie( "Flags" )
g_auth_bGood = FALSE
g_auth_bRemember = FALSE

IF g_auth_bAccept THEN
	IF 0 < g_auth_nUserID THEN
		g_auth_bGood = TRUE
	END IF
END IF
IF g_auth_bGood THEN
	g_auth_bRemember = authCookieBool( "Remember" )
	IF g_auth_bRemember THEN
		authSetAllCookies
	END IF
ELSE
	Response.Redirect "logon.asp"
END IF





%>