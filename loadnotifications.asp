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
<!--#include file="scripts\include_pagelist.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<!--#include file="scripts\remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts\include_announce.asp"-->
<!--#include file="scripts\include_rssannounce.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\rss.asp"-->
<!--#include file="scripts\include_forum.asp"-->
<!--#include file="scripts\rssweather.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_emailnotificationformat.asp"-->
<!--#include file="scripts/include_pagebody.asp"-->
<!--#include file="scripts/include_picture2.asp"-->
<!--#include file="scripts\include_emailnotification.asp"-->
<%


FUNCTION genEmail()

	DIM	s
	s = ""

	IF "localhost" <> LCASE(Request.ServerVariables("SERVER_NAME")) THEN
		IF needsNotification() THEN
			'Response.Write "<p>Needs Notification</p>" & vbCRLF
			'writeNotificationDate
			s = processNotification()
			IF "" <> s THEN
				s = "Emails delivered: " & s
			ELSE
				s = "x"
			END IF
		ELSE
			s = ""
		END IF
	ELSE
		s = "localhost will not send emails"
	END IF

	genEmail = s

END FUNCTION


FUNCTION getEmail()

	DIM	sData
	sData = ""
	DIM	sCacheFile
	sCacheFile = "emailnotice.htm"
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, "h", 24, "d" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sData = genEmail()
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", sCacheFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		ELSE
			cache_deleteFile "site", sCacheFile
		END IF
	END IF
	getEmail = sData
	
END FUNCTION


DIM	sData
sData = getEmail()
IF 0 < LEN(sData) THEN
	Response.ContentType = "text/text"
	Response.Write sData
ELSE
	Response.Status = "404 Not Found"
END IF

Response.Flush
Response.End


%>

<!--#include file="scripts\include_db_end.asp"-->
