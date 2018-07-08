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
<!--#include file="scripts/remind_utils.asp"-->
<!--#include file="scripts/include_navtabs.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
Response.Write "<title>" & Server.HTMLEncode("Registration - " & g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Registration - Notification</title>
<%
END IF
%>

<meta name="navigate" content="calendar-register">
<meta name="navtitle" content="Email Subscriptions">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="shortcut icon" href="./favicon.ico">
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
</head>

<body>
<%
DIM g_sNavigateTabLabel
g_sNavigateTabLabel = "(home|calendar-register)"

%>
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="config\header.asp"-->

<%

DIM g_subscriptions_sDataFolder

FUNCTION findSubscriptionsFolder()
	DIM	sFolder
	
	IF "" = g_subscriptions_sDataFolder THEN
		sFolder = findAppDataFolder( "emailsubscriptions" )
		IF "" <> sFolder THEN
			findSubscriptionsFolder = sFolder
			g_subscriptions_sDataFolder = sFolder
		ELSE
			findSubscriptionsFolder = ""
		END IF
	ELSE
		findSubscriptionsFolder = g_subscriptions_sDataFolder
	END IF
END FUNCTION


FUNCTION findSubscriptionsFile()

	findSubscriptionsFile = ""
	DIM	sFolder
	DIM	sFile

	sFile = ""
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath(sFolder,"subscriptions.txt")
		IF g_oFSO.FileExists(sFile) THEN
			findSubscriptionsFile = sFile
		END IF
	END IF

END FUNCTION


FUNCTION getTempFilename()

	DIM	sFolder
	sFolder = findSubscriptionsFolder()
	
	DIM	sTemp
	sTemp = Session.SessionID & "_" & HOUR(NOW) & "_" & MINUTE(NOW) & "_" & SECOND(NOW) & ".txt"

	getTempFilename = g_oFSO.BuildPath(sFolder,sTemp)

END FUNCTION



SUB copySubscriptionsWOemail( sEMail, oTFile, oSFile )

	DIM	sEMailLC
	sEMailLC = LCASE(sEMail)
	
	DIM	sLine
	
	DO UNTIL oSFile.AtEndOfStream
		sLine = TRIM(oSFile.ReadLine)
		IF "" <> sLine THEN
			IF LCASE(sLine) <> sEMailLC THEN
				oTFile.WriteLine sLine
			END IF
		END IF
	LOOP

END SUB



SUB message( sText )
%>
<p><%=sText%> is complete</p>
<%
END SUB




SUB listServMail( sCommand, sEmail, sList )




	DIM oSMTP

	SET oSMTP = mailMakeSMTP()
	IF NOT Nothing IS oSMTP THEN
		oSMTP.Server = mailServer()
		
		DIM	sAccount
		DIM	sPW
		
		sAccount = g_sAnnonUser
		sPW = g_sAnnonPW
		
		'sAccount = ""
		'sPW = ""
		
		IF "" <> sAccount  AND  "" <> sPW THEN
			oSMTP.User = sAccount
			oSMTP.PW = sPW
		END IF
		
		
		oSMTP.From = sEmail
		oSMTP.ReplyTo = sEmail
        oSMTP.DeliveryMethod = k_nMailDeliveryDefault
		oSMTP.AddRecipient LCASE(sList & "-" & sCommand) & "@" & g_sDomain
		oSMTP.Subject = sCommand & " " & sList
		oSMTP.Body = sCommand & " " & sList & vbCRLF
				
		'oSMTP.From = sAccount
		'oSMTP.AddRecipient sList & "-" & sCommand & "-" & REPLACE(sEmail,"@","=") & "@" & g_sDomain
		'oSMTP.Subject = sCommand & " " & sList
		'oSMTP.Body = sCommand & " " & sList & vbCRLF
				

		'send it
		DIM	sError
		sError = oSMTP.Send
		IF "0:" = LEFT(sError,2) THEN
			Response.Write "<p>Request sent to list</p>"
		ELSE
			Response.Write "<p>Request Failure: " & sError & "</p>"
		END IF


	END IF



END SUB




SUB register( sEMail, sList )

	DIM	oTFile
	DIM	sTFile
	DIM	oSFile
	DIM	sSFile
	
	IF FALSE THEN
	IF "UpcomingEvents" = sList THEN
	
		sTFile = getTempFilename()
		sSFile = findSubscriptionsFile()
		SET oTFile = g_oFSO.CreateTextFile( sTFile, TRUE )
		SET oSFile = g_oFSO.OpenTextFile( sSFile, 1 )
		
		copySubscriptionsWOemail sEMail, oTFile, oSFile
		
		oSFile.Close
		SET oSFile = Nothing
		
		oTFile.WriteLine sEMail
		oTFile.Close
		SET oTFile = Nothing
	
		IF g_oFSO.FileExists( sSFile ) THEN g_oFSO.DeleteFile sSFile, TRUE
		g_oFSO.CopyFile sTFile, sSFile, TRUE
		g_oFSO.DeleteFile sTFile, TRUE
	
	END IF
	END IF
	
	listServMail "Subscribe", sEMail, sList

%>
<p>An email will be sent to you to confirm your registration from <span class="pobox"><%=sList%>
-help</span>. If you have a spam checker enabled please make sure that email 
address is allowed.&nbsp; The registration is not complete until you respond to 
the email.</p>
<%
	
	'message "Registration"
	
END SUB



SUB unregister( sEMail, sList )

	DIM	oTFile
	DIM	sTFile
	DIM	oSFile
	DIM	sSFile
	
	IF FALSE THEN
	IF "UpcomingEvents" = sList THEN
	
		sTFile = getTempFilename()
		sSFile = findSubscriptionsFile()
		SET oTFile = g_oFSO.CreateTextFile( sTFile, TRUE )
		SET oSFile = g_oFSO.OpenTextFile( sSFile, 1 )
		
		copySubscriptionsWOemail sEMail, oTFile, oSFile
		
		oSFile.Close
		SET oSFile = Nothing
		
		oTFile.Close
		SET oTFile = Nothing
	
		IF g_oFSO.FileExists( sSFile ) THEN g_oFSO.DeleteFile sSFile, TRUE
		g_oFSO.CopyFile sTFile, sSFile, TRUE
		g_oFSO.DeleteFile sTFile, TRUE
	
	END IF
	END IF

	listServMail "unsubscribe", sEMail, sList

%>
<p>An email will be sent to you to confirm your removal from the list from <span class="pobox"><%=sList%>
-help</span>. If you have a spam checker enabled please make sure that email 
address is allowed.&nbsp; Removal will not be completed until you respond to the 
email.</p>
<%



	'message "Unregistration"
	
END SUB




DIM	sEMail
DIM	sAction
DIM	sList

sEMail = Request("email")
sAction = Request("action")
sList = Request("list")

'NewsletterNotice
'UpcomingEvents

IF "" = sList THEN
	sList = "UpcomingEvents"
END IF


SELECT CASE sAction
CASE "register"
	register sEMail, sList
CASE "unregister"
	unregister sEMail, sList
END SELECT


%>
<!--#include file="scripts\page_end.asp"-->
<!--webbot bot="Include" u-include="_private/byline.htm" tag="BODY" startspan -->

<script language="JavaScript">
<!--

function makeByLine()
{
	document.write( '<' + 'script language="JavaScript" src="http://' );
	if ( "localhost" == location.hostname )
	{
		document.writeln( 'localhost/BearConsultingGroup' );
	}
	else
	{
		document.writeln( 'BearConsultingGroup.com' );
	}
	document.writeln( '/designby_small.js"></' + 'script>' );
}
makeByLine()

//-->
</script>
<!--script language="JavaScript" src="http://BearConsultingGroup.com/designbyadvert.js"></script-->

<!--webbot bot="Include" i-checksum="20551" endspan -->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
<%
%>