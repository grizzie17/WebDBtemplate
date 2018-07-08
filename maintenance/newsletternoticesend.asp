<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit






%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\htmlformat.asp"-->
<!--#include file="scripts\mailserver.asp"-->
<!--#include file="scripts\mailsend.asp"-->
<%


FUNCTION recConfigString( oRS, sField )
	DIM	oField
	SET oField = oRS.Fields(sField)
	DIM	sTemp
	sTemp = recString( oField )
	IF "" <> sTemp THEN
		recConfigString = sTemp
	ELSE
		recConfigString = ""
	END IF
	
END FUNCTION



'DIM	g_sDomain
'DIM	g_sSiteTitle
'DIM	g_sSiteChapter
'DIM	g_sMailServer
'DIM	g_sAnnonUser
'DIM	g_sAnnonPW

DIM	sSelect
sSelect = "" _
	&	"SELECT " _
	&	"	* " _
	&	"FROM " _
	&	"	Config " _
	&	";"

DIM	oRS
SET oRS = dbQueryRead( g_DC, sSelect )

IF NOT oRS IS Nothing THEN

	IF NOT oRS.EOF THEN
	
		g_sDomain = recConfigString( oRS, "Domain" )
		g_sSiteTitle = recConfigString( oRS, "SiteTitle" )
		g_sSiteChapter = recConfigString( oRS, "SiteChapter" )
		g_sMailServer = recConfigString( oRS, "MailServer" )
		g_sAnnonUser = recConfigString( oRS, "MailUser" )
		g_sAnnonPW = recConfigString( oRS, "MailPW" )

	END IF
	
END IF

DIM q_sMonth
q_sMonth = Request("Month")

DIM q_sComments
q_sComments = Request("Comments")


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<title>Untitled 1</title>
</head>

<body>

<%


FUNCTION setupSMTP()
	SET setupSMTP = Nothing
	DIM oSMTP

	SET oSMTP = mailMakeSMTP()
	IF NOT Nothing IS oSMTP THEN
		oSMTP.Server = mailServer()
		
		DIM	sAccount
		DIM	sPW
		
		sAccount = g_sAnnonUser
		sPW = g_sAnnonPW
		
		IF "" <> sAccount  AND  "" <> sPW THEN
			oSMTP.User = sAccount
			oSMTP.PW = sPW
		END IF
		
		
		
		oSMTP.From = g_sAnnonUser
		IF "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
			oSMTP.FromName = "Newsletter Notification"
			oSMTP.AddRecipient2 "NewsletterNotice@" & g_sDomain, "Newsletter Notice"
		ELSE
			oSMTP.FromName = "Local " & g_sDomain
			oSMTP.AddRecipientBCC "Webmaster@" & g_sDomain
		END IF

		SET setupSMTP = oSMTP
		SET oSMTP = Nothing
	END IF
END FUNCTION


FUNCTION makeMailBody()
	DIM	s
	s = "The " & Server.HTMLEncode(g_sSiteTitle) & " (" & g_sSiteChapter & ") Newsletter " _
		& "for " & MonthName(q_sMonth) & " is now available."
	DIM	sURLNews
	sURLNews = "www." & g_sDomain & "/Newsletter"
	DIM	sURL
	sURL = "http://" & sURLNews
	DIM	sHTX
	IF "" <> q_sComments THEN
		sHTX = HTMLFormatCRLF( q_sComments )
	ELSE
		sHTX = ""
	END IF
	DIM sBody
	sBody = "" _
		&	"<html>" & vbCRLF _
		&	"<body>" & vbCRLF _
		&	"<p>" & s & "</p>" & vbCRLF _
		&	"<p><a target=""_blank"" href=""" & sURL & """>" & sURLNews & "</a></p>" & vbCRLF _
		&	sHTX & vbCRLF _
		&	"<p style=""color:gray"">Do not reply to this email.<br>" _
		&	"List edits must be performed on the website.</p>" & vbCRLF _
		&	"</body>" & vbCRLF _
		&	"</html>" & vbCRLF
	makeMailBody = sBody
END FUNCTION


DIM	oSMTP

SET oSMTP = setupSMTP()
IF NOT oSMTP IS Nothing THEN
	DIM	sBody
	sBody = makeMailBody()
	
	oSMTP.Subject = "GWRRA: " & g_sSiteTitle & " - " & MonthName(q_sMonth) & " Newsletter"
	oSMTP.Body = sBody

	DIM	sError
	sError = oSMTP.Send
	IF "0:" = LEFT(sError,2) THEN
		Response.Write "<p>Mail successfully sent</p>"
	ELSE
		Response.Write "<p>Problem sending email</p>"
		Response.Write "<p>" & sError & "</p>"
	END IF
	SET oSMTP = Nothing

END IF






%>




</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
