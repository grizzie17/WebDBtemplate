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
<!--#include file="scripts\mailserver.asp"-->
<!--#include file="scripts\mailsend.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Calendar Registration Request</title>
<link rel="shortcut icon" href="./favicon.ico">
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
</head>

<body>
<%
DIM g_sNavigateTabLabel
g_sNavigateTabLabel = "(home|register-request)"

%>
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="app_data\layout\header.asp"-->

<%

DIM	sEmail
DIM	sList

sEmail = Request("email")
sList = Request("list")

IF "" <> sEmail THEN

	
	DIM oSMTP

	SET oSMTP = mailMakeSMTP()
	IF NOT Nothing IS oSMTP THEN
		'oSMTP.Server = "mail.rocketcitywings.org"
		oSMTP.Server = mailServer()
		
		DIM	sAccount
		DIM	sPW
		
		sAccount = g_sAnnonUser
		sPW = g_sAnnonPW
		
		IF "" <> sAccount  AND  "" <> sPW THEN
			oSMTP.User = sAccount
			oSMTP.PW = sPW
		END IF
		
		DIM sListAction
		
		SELECT CASE sList
		CASE "UpcomingEvents"
			sListAction = "Upcoming Events"
		CASE "NewsletterNotice"
			sListAction = "Newsletter and Event "
		CASE ELSE
			sListAction = "Upcoming Events"
		END SELECT

		oSMTP.From = g_sAnnonUser
		oSMTP.FromName = "Register Notification"
	
		oSMTP.AddRecipient sEMail
		
		oSMTP.Subject = g_sSiteName & " Calendar registration"
			
		oSMTP.Body = "" _
			&	"<html>" & vbCRLF _
			&	"<head>" & vbCRLF _
			&	"</head>" & vbCRLF _
			&	"<body>" & vbCRLF _
			&	"<p>" & sListAction & " " & Server.HTMLEncode(g_sSiteName) & " - Email Registration</p>" & vbCRLF _
			&	"<p>To register: " _
			&		"<a href=""" _
			&		"http://www." & g_sDomain & "/notification.asp?action=register&list=" & sList & "&email=" & Server.URLEncode(sEMail) _
			&		""">" _
			&		"Click Here" _
			&		"</a>" _
			&	"</p>" & vbCRLF _
			&	"<p>To unregister: " _
			&		"<a href=""" _
			&		"http://www." & g_sDomain & "/notification.asp?action=unregister&list=" & sList & "&email=" & Server.URLEncode(sEMail) _
			&		""">" _
			&		"Click Here" _
			&		"</a>" _
			&	"</p>" & vbCRLF _
			&	"<p>Do not reply to this email</p>" & vbCRLF _
			&	"</body>" & vbCRLF _
			&	"</html>" & vbCRLF
		
		
		'send it
		DIM	sError
		sError = oSMTP.Send
		IF "0:" = LEFT(sError,2) THEN
%>
<p>An email has been sent to <%=Server.HTMLEncode(sEmail)%>.&nbsp; <br>
You must respond to the email by clicking on either the register or unregister 
link.</p>
<p>If you have a spam program you must add &quot;<font color="#0000FF"><%=g_sAnnonUser%></font>&quot; 
to your approved list.</p>
<%
		ELSE
%>
<p>An error occurred attempting to send the registration email. <%=sError%></p>
<%
		END IF
	END IF
	
	SET oSMTP = Nothing

END IF

%>

<p><a href="calendar.asp">Return to Calendar Page</a></p>
<!--#include file="scripts\page_end.asp"-->
<!--webbot bot="Include" u-include="_private/byline.htm" tag="BODY" -->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
<%
%>