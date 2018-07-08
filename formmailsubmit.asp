<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT




DIM	sPageTitle
sPageTitle = Request.Form("mailPageTitle")
IF "" = sPageTitle THEN
	sPageTitle = "Comment Submission"
END IF



%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\htmlformat.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\include_guid.asp"-->
<!--#include file="scripts\mailserver.asp"-->
<!--#include file="scripts\mailsend.asp"-->
<%





DIM g_sCommentKey
DIM g_sSessionGUID

g_sCommentKey = g_sCookiePrefix & "_SendMessage"

g_sSessionGUID = getGuidChars(Session(g_sCommentKey))



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<%
Response.Write "<title>" & sPageTitle & " - " & Server.HTMLEncode(g_sSiteName) & "</title>" & vbCRLF
IF FALSE THEN
%>
<title>Form Mail Submit</title>
<%
END IF
%>
<link rel="shortcut icon" href="favicon.ico">
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
</head>

<body>
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="config\header.asp"-->
<%



FUNCTION processFieldList( sOrder )
	DIM	sMessage
	DIM	sField
	DIM	sValue
	DIM	tsOrder
	tsOrder = "," & sOrder & ","
	
	sMessage = "<table>" & vbCRLF
	IF "" <> sOrder THEN
		DIM	aFields
		aFields = SPLIT(sOrder,",")
		FOR EACH sField in aFields
			sValue = Request.Form(sField)
			IF "" <> sValue THEN
				sMessage = sMessage _
					& "<tr>" _
					& "<th align=""right"" valign=""top"">" _
					& Server.HTMLEncode( sField ) _
					& "&nbsp;</th>" _
					& "<td valign=""top"">" _
					& htmlFormatCRLF( sValue ) _
					& "</td>" _
					& "</tr>" & vbCRLF
			END IF
		NEXT 'sField
	END IF
	FOR EACH sField In Request.Form
		IF "mail" <> LEFT( sField, 4 ) THEN
			IF 0 = INSTR( tsOrder, ","&sField&",") THEN
				sValue = Request.Form(sField)
				IF "" <> sValue THEN
					sMessage = sMessage _
						& "<tr>" _
						& "<th align=""right"" valign=""top"">" _
						& Server.HTMLEncode( sField ) _
						& "&nbsp;</th>" _
						& "<td valign=""top"">" _
						& htmlFormatCRLF( sValue ) _
						& "</td>" _
						& "</tr>" & vbCRLF
				END IF
			END IF
		END IF
	Next 'sField
	processFieldList = sMessage & "</table>" & vbCRLF
	
END FUNCTION


FUNCTION checkRequiredFields( sFields )

	checkRequiredFields = TRUE
	DIM	aFields
	DIM	sField
	DIM	sValue
	
	IF "" <> sFields THEN
		aFields = SPLIT( sFields, "," )
		FOR EACH sField IN aFields
			sValue = Request.Form(sField)
			IF "" = sValue THEN
				checkRequiredFields = FALSE
				Response.Write "Required field '" & sField & "' is missing<br>" & vbCRLF
			END IF
		NEXT 'sField
	END IF
END FUNCTION


FUNCTION processAddrDomain( sAddr )
	DIM aTemp
	DIM	sTemp
	DIM	sList
	aTemp = SPLIT( sAddr, ";" )
	sList = ""
	FOR EACH sTemp IN aTemp
		sTemp = TRIM(sTemp)
		IF "" <> sTemp THEN
			sTemp = REPLACE(sTemp, "*", "@" )
			sTemp = REPLACE(sTemp, ":", "." )
			DO WHILE 0 < INSTR(sTemp, " ")
				sTemp = REPLACE(sTemp, " ", "" )
			LOOP
			IF 0 = INSTR(sTemp,"@") THEN sTemp = sTemp & "@" & g_sDomain
			sList = sList & ";" & sTemp
		END IF
	NEXT 'sTemp
	processAddrDomain = MID(sList,2)
END FUNCTION




DIM	sField
DIM	sValue
DIM	sMessage
DIM	oSMTP


	DIM	sTo
	DIM	sFrom
	DIM	sFromAddr
	DIM	sFromName
	DIM	sCC
	DIM	sBCC
	DIM	sReplyTo
	DIM	sSubject
	DIM	sPriority
	DIM	sRedirect
	DIM	sSendSuccess
	
	DIM	sMailGUID
	
	DIM	sRequiredFields
	DIM	sOrderFields
	
	sMailGUID = Request.Form("mailSessionGUID")
	
	sTo = Request.Form("mailTo")
	sFrom = Request.Form("mailFrom")
	sFromAddr = Request.Form("mailFromAddr")
	sFromName = Request.Form("mailFromName")
	sCC = Request.Form("mailCC")
	sBCC = Request.Form("mailBCC")
	sReplyTo = Request.Form("mailReplyTo")
	sSubject = Request.Form("mailSubject")
	sPriority = Request.Form("mailPriority")
	sRedirect = Request.Form("mailRedirect")
	sSendSuccess = Request.Form("mailSendSuccess")
	
	sRequiredFields = Request.Form("mailRequired")
	sOrderFields = Request.Form("mailOrderFields")
	
	IF "" <> sTo THEN
		sTo = processAddrDomain( sTo )
	END IF
	IF "" <> sCC THEN
		sCC = processAddrDomain( sCC )
	END IF
	IF "" <> sBCC THEN
		sBCC = processAddrDomain( sBCC )
	END IF
	
	IF "" <> sFromAddr THEN
		sFromAddr = processAddrDomain( sFromAddr )
	END IF
	IF "" = sFrom THEN
		IF 0 < LEN(sFromName) THEN
			sFrom = sFromName & "[" & sFromAddr & "]"
		ELSE
			sFrom = sFromAddr
			sFromName = sFrom
		END IF
	ELSE
		IF "" = sFromName THEN
			sFromName = sFrom
		END IF
	END IF
	
	IF "" = sReplyTo THEN
		sReplyTo = sFrom
	END IF
	
	sMessage = processFieldList( sOrderFields )
	
	DIM	sHtmlMessage
	sHtmlMessage =  "" _
			&	"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCRLF _
			&	"<html>" & vbCRLF _
			&	"<head>" & vbCRLF _
			&	"<meta http-equiv=""Content-Language"" content=""en-us"">" & vbCRLF _
			&	"<title>" & Server.HTMLEncode(sSubject) & "</title>" & vbCRLF _
			&	"</head>" & vbCRLF _
			&	"<body bgcolor=""#FFFFFF"">" & vbCRLF _
			&	"<h2>" & Server.HTMLEncode(sSubject) & "</h2>" & vbCRLF _
			&	"<h3>" & Server.HTMLEncode(sFromName) & "</h3>" & vbCRLF _
			&	sMessage _
			&	"</body>" & vbCRLF _
			&	"</html>" & vbCRLF


	IF CSTR(sMailGUID) = CSTR(g_sSessionGUID)  _
			AND  "" <> sTo  AND  "" <> sFrom  AND  "" <> sSubject  _
			AND  checkRequiredFields(sRequiredFields) THEN

		'Response.Write sHtmlMessage
		
		DIM	sStatus
		DIM	sError
	
		SET oSMTP = mailMakeSMTP()

		IF g_bMailPickup THEN
			oSMTP.DeliveryMethod = k_nMailDeliveryPickup
		END IF

		oSMTP.Server = mailServer()
		oSMTP.User = g_sAnnonUser
		oSMTP.PW = g_sAnnonPW


		mailLoadNames oSMTP, sFrom, sReplyTo, sTo, sCC, sBCC
			
		oSMTP.Subject = sSubject
			
		oSMTP.Body = sHtmlMessage
		sError = oSMTP.Send


		'nStatus = mailsend( sFrom, sReplyTo, sTo, sCC, sBCC, sSubject, sHtmlMessage )
	
		IF "0:" <> LEFT(sError,2) THEN
		%>
<p><span style="color:red; font-size:xx-large;">Oops! </span><span style="font-size:large">An error occurred during the submission of your comments.</span></p>
		<%
			Response.Write "<p>" & Server.HTMLEncode(sError) & "</p>" & vbCRLF

		ELSE
			IF "" <> sSendSuccess THEN
				Response.Write "<p>" & Server.HTMLEncode(sSendSuccess) & "</p>" & vbCRLF
			ELSE
		%>
<p>Thank you for your comments. ... Your message has been successfully sent.</p>
			<%
			END IF
			IF "" <> sRedirect THEN
			%>
<script language="JavaScript" type="text/javascript">
<!--
var sTargetURL = "<%=sRedirect%>";

function doRedirect()
{
    setTimeout( "window.location.href = sTargetURL", 1*1000 );
}

//-->
</script>

<script language="JavaScript1.1" type="text/javascript">
<!--
function doRedirect()
{
    window.location.replace( sTargetURL );
}

doRedirect();

//-->
</script>
			<%
			END IF
		END IF
	ELSEIF "" <> CSTR(sMailGUID)  AND  CSTR(sMailGUID) <> CSTR(g_sSessionGUID) THEN
	%>
<p><span style="color:red; font-size:x-large">Unrecognized session ... </span>
<span style="font-size: large">Message has NOT been sent.</span></p>
	<%
		'Response.Write "passed = " & sMailGUID & " (" & LEN(sMailGUID) & ")<br>"
		'Response.Write "session = " & g_sSessionGUID & " (" & LEN(g_sSessionGUID) & ")<br>"
		
	ELSE
	%>
<p><span style="color:red; font-size:x-large">Missing required fields ... </span>
<span style="font-size: large">Message has NOT been sent.</span></p>
	<%

	END IF


%>
<!--#include file="scripts\page_end.asp"-->

</body>
</html>
<!--#include file="scripts\include_db_begin.asp"-->
