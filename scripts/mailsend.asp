<%

CONST k_nMailPriorityNone = 0
CONST k_nMailPriorityLow = 1
CONST k_nMailPriorityMedium = 2
CONST k_nMailPriorityHigh = 3


CONST k_nMailDeliveryDefault = 0
CONST k_nMailDeliveryPickup = 1
CONST k_nMailDeliveryPort = 2

FUNCTION mailPriority( n )
	DIM	j
	IF ISNUMERIC(n) THEN
		j = CLNG(n)
		IF j < k_nMailPriorityNone  OR  k_nMailPriorityHigh < j THEN j = k_nMailPriorityNone
	ELSE
		j = k_nMailPriorityNone
	END IF
	mailPriority = j
END FUNCTION



' function to look at the string 's' and parse it to
' retrieve the name and address parts
SUB parseAddr( sName, sAddr, sInput )
	DIM	i
	DIM	s
	
	s = sInput
	
	i = INSTR(1,s,"[",vbTextCompare)
	IF 0 < i THEN
		sAddr = MID(s,i)
		sAddr = REPLACE(sAddr,"[","",1,-1,vbTextCompare)
		sAddr = TRIM(REPLACE(sAddr,"]","",1,-1,vbTextCompare))
		sName = TRIM(LEFT(s,i-1))
	ELSE
		i = INSTR(1,s,"<",vbTextCompare)
		IF 0 < i THEN
			sAddr = MID(s,i)
			sAddr = REPLACE(sAddr,"<","",1,-1,vbTextCompare)
			sAddr = TRIM(REPLACE(sAddr,">","",1,-1,vbTextCompare))
			sName = TRIM(LEFT(s,i-1))
		ELSE
			sAddr = TRIM(s)
			sName = ""
		END IF
	END IF
	IF 0 < LEN(sName) THEN
		sName = TRIM(REPLACE(sName,"""","",1,-1,vbTextCompare))
	END IF
	IF 0 = INSTR(1,sAddr,"@",vbTextCompare)  AND  0 < INSTR(1,sName,"@",vbTextCompare) THEN
		s = sName
		sName = sAddr
		sAddr = s
	END IF
END SUB






CLASS PersistsSMTP

	PRIVATE	m_oSMTP
	
	PRIVATE m_nInline
	
	SUB Class_Initialize()
		SET m_oSMTP = Nothing
		m_nInline = 0
		'Response.Write "Class_Initialize: Persists<br>"
		'ON ERROR Resume Next
		'SET m_oSMTP = Server.CreateObject("Persits.MailSender")
		'ON ERROR Goto 0
	END SUB
	
	SUB Class_Terminate()
		'Response.Write "Class_Terminate: Persists<br>"
		SET m_oSMTP = Nothing
	END SUB

	PROPERTY SET SMTP( o )
		SET m_oSMTP = o
	END PROPERTY

	PUBLIC DEFAULT PROPERTY GET SMTP
		'IF Nothing IS m_oSMTP THEN Response.Write "Problem creating component Persits.MailSender<br>"
		SET SMTP = m_oSMTP
	END PROPERTY

	PROPERTY LET Server( s )
		m_oSMTP.Host = s
	END PROPERTY
	
	PROPERTY LET DeliveryMethod( n )
	END PROPERTY
	
	PROPERTY LET User( s )
		ON ERROR Resume Next
		m_oSMTP.Username = s
		ON ERROR Goto 0
	END PROPERTY
	
	PROPERTY LET PW( s )
		ON ERROR Resume Next
		m_oSMTP.Password = s
		ON ERROR Goto 0
	END PROPERTY
	
	PROPERTY LET From( s )
		m_oSMTP.From = s
	END PROPERTY
	
	PROPERTY LET FromName( s )
		m_oSMTP.FromName = s
	END PROPERTY
	
	PROPERTY LET ReplyTo( s )
		m_oSMTP.AddReplyTo s
	END PROPERTY
	
	SUB AddRecipient( sAddr )
		m_oSMTP.AddAddress sAddr
	END SUB
	
	SUB AddRecipient2( sAddr, sName )
		m_oSMTP.AddAddress sAddr, sName
	END SUB
	
	SUB AddRecipientCC( sAddr )
		m_oSMTP.AddCC sAddr
	END SUB
	
	SUB AddRecipientCC2( sAddr, sName )
		m_oSMTP.AddCC sAddr, sName
	END SUB
	
	SUB AddRecipientBCC( sAddr )
		m_oSMTP.AddBCC sAddr
	END SUB
	
	SUB AddRecipientBCC2( sAddr, sName )
		m_oSMTP.AddBCC sAddr, sName
	END SUB

	FUNCTION AddInlineAttachment( sFile )
		DIM	sAlias
		DIM	sName
		DIM	i
		i = INSTRREV( sFile, "\" )
		IF 0 < i THEN
			sName = MID( sFile, i+1 )
		ELSE
			sName = sFile
		END IF
		m_nInline = m_nInline + 1
		sAlias = "inline" & m_nInline & "_" & sName
		ON ERROR Resume Next
		m_oSMTP.AddEmbeddedImage sFile, sAlias
		IF Err <> 0 THEN
			AddInlineAttachment = ""
		ELSE
			AddInlineAttachment = sAlias
		END IF
		ON ERROR Goto 0
	END FUNCTION
	
	SUB AddAttachment( sFile )
		m_oSMTP.AddAttachment sFile
	END SUB
	

	PROPERTY LET Subject( s )
		m_oSMTP.Subject = s
	END PROPERTY
	
	PROPERTY LET Body( s )
		m_oSMTP.Body = s
		IF "<html>" = LCASE(LEFT(s,6))  _
				OR  "<!doctype html" = LCASE(LEFT(s,14)) THEN
			m_oSMTP.IsHTML = TRUE
		END IF
	END PROPERTY

	PROPERTY LET Priority( n )
		DIM	p
		SELECT CASE mailPriority( n )
		CASE k_nMailPriorityNone
			p = 0
		CASE k_nMailPriorityLow
			p = 5
		CASE k_nMailPriorityMedium
			p = 3
		CASE k_nMailPriorityHigh
			p = 1
		CASE ELSE
			p = 0
		END SELECT
		m_oSMTP.Priority = p
	END PROPERTY

	FUNCTION Send()
		ON ERROR Resume Next
		IF m_oSMTP.Send() THEN
			Send = "0:Success"
		ELSE
			Send = Err.Number & ":" & Err.Description
		END IF
		ON ERROR Goto 0
	END FUNCTION

END CLASS


CLASS JMailSMTP

	PRIVATE	m_oSMTP
	PRIVATE	m_sServer
	
	SUB Class_Initialize()
		SET m_oSMTP = Nothing
		m_sServer = ""
		'Response.Write "Class_Initialize: JMail<br>"
		'ON ERROR Resume Next
		'SET m_oSMTP = Server.CreateObject("JMail.Message")
		'ON ERROR Goto 0
	END SUB
	
	SUB Class_Terminate()
		'Response.Write "Class_Terminate: JMail<br>"
		SET m_oSMTP = Nothing
	END SUB

	PUBLIC PROPERTY GET SMTP
		'IF Nothing IS m_oJMail THEN Response.Write "Problem creating component JMail.Message<br>"
		SET SMTP = m_oSMTP
	END PROPERTY

	PROPERTY SET SMTP( o )
		SET m_oSMTP = o
		m_oSMTP.ISOEncodeHeaders = FALSE
	END PROPERTY

	PROPERTY LET Server( s )
		m_sServer = s
	END PROPERTY

	PROPERTY LET DeliveryMethod( n )
	END PROPERTY

	PROPERTY LET User( s )
		m_oSMTP.MailServerUserName = s
	END PROPERTY
	
	PROPERTY LET PW( s )
		m_oSMTP.MailServerPassword = s
	END PROPERTY
	
	PROPERTY LET From( s )
		m_oSMTP.From = s
	END PROPERTY
	
	PROPERTY LET FromName( s )
		m_oSMTP.FromName = s
	END PROPERTY
	
	PROPERTY LET ReplyTo( s )
		m_oSMTP.ReplyTo = s
	END PROPERTY
	
	SUB AddRecipient( sAddr )
		m_oSMTP.AddRecipient sAddr
	END SUB
	
	SUB AddRecipient2( sAddr, sName )
		m_oSMTP.AddRecipient sAddr, sName
	END SUB
	
	SUB AddRecipientCC( sAddr )
		m_oSMTP.AddRecipientCC sAddr
	END SUB
	
	SUB AddRecipientCC2( sAddr, sName )
		m_oSMTP.AddRecipientCC sAddr, sName
	END SUB
	
	SUB AddRecipientBCC( sAddr )
		m_oSMTP.AddRecipientBCC sAddr
	END SUB
	
	SUB AddRecipientBCC2( sAddr, sName )
		m_oSMTP.AddRecipientBCC sAddr, sName
	END SUB
	
	FUNCTION AddInlineAttachment( sFile )
		AddInlineAttachment = m_oSMTP.AddAttachment( sFile, TRUE )
	END FUNCTION
	
	SUB AddAttachment( sFile )
		m_oSMTP.AddAttachment sFile, FALSE
	END SUB
	
	PROPERTY LET Subject( s )
		m_oSMTP.Subject = s
	END PROPERTY
	
	PROPERTY LET Body( s )
		IF "<html>" = LCASE(LEFT(s,6))  _
				OR  "<!doctype html" = LCASE(LEFT(s,14)) THEN
			m_oSMTP.HTMLBody = s
		ELSE
			m_oSMTP.Body = s
		END IF
	END PROPERTY

	PROPERTY LET Priority( n )
		DIM	p
		SELECT CASE mailPriority( n )
		CASE k_nMailPriorityNone
			p = 3
		CASE k_nMailPriorityLow
			p = 5
		CASE k_nMailPriorityMedium
			p = 3
		CASE k_nMailPriorityHigh
			p = 1
		CASE ELSE
			p = 3
		END SELECT
		m_oSMTP.Priority = p
	END PROPERTY

	FUNCTION Send()
		m_oSMTP.Silent = TRUE
		IF m_oSMTP.Send( m_sServer ) THEN
			Send = "0:Success"
		ELSE
			Send = m_oSMTP.ErrorCode & ":" & m_oSMTP.ErrorMEssage
		END IF
	END FUNCTION

END CLASS


CLASS CDONTSSMTP

	PRIVATE	m_oSMTP
	PRIVATE	m_sServer
	
	PRIVATE	m_sFROM
	PRIVATE m_sFromName
	PRIVATE	m_sTO
	PRIVATE	m_sCC
	PRIVATE	m_sBCC
	
	PRIVATE	m_nInline
	PRIVATE m_aAttachments()
	PRIVATE m_nAttachments

	SUB Class_Initialize()
		SET m_oSMTP = Nothing
		m_sServer = ""
		m_sFROM = ""
		m_sFromName = ""
		m_sTO = ""
		m_sCC = ""
		m_sBCC = ""
		m_nInline = 0
		REDIM m_aAttachments(5)
		m_nAttachments = 0
		'Response.Write "Class_Initialize: CDONTS<br>"
		'ON ERROR Resume Next
		'SET m_oSMTP = Server.CreateObject("CDONTS.NewMail")
		'ON ERROR Goto 0
	END SUB
	
	SUB Class_Terminate()
		'Response.Write "Class_Terminate: CDONTS<br>"
		SET m_oSMTP = Nothing
	END SUB

	PRIVATE SUB addAttachmentList( s )
		IF UBOUND(m_aAttachments) <= m_nAttachments+1 THEN
			REDIM PRESERVE m_aAttachments(UBOUND(m_aAttachments)+5)
		END IF
		m_aAttachments(m_nAttachments) = s
		m_nAttachments = m_nAttachments + 1
	END SUB

	PRIVATE SUB addAttachmentFile( sFile )
		DIM	sName
		DIM	i
		i = INSTRREV( sFile, "\" )
		IF 0 < i THEN
			sName = MID( sFile, i+1 )
		ELSE
			sName = sFile
		END IF
		ON ERROR Resume Next
		m_oSMTP.AttachFile sFile, sName
		ON ERROR Goto 0
	END SUB
	

	PUBLIC PROPERTY GET SMTP
		'IF Nothing IS m_oSMTP THEN Response.Write "Problem creating component CDONTS.NewMail<br>"
		SET SMTP = m_oSMTP
	END PROPERTY

	PROPERTY SET SMTP( o )
		SET m_oSMTP = o
	END PROPERTY

	PROPERTY LET Server( s )
		m_sServer = s
	END PROPERTY

	PROPERTY LET DeliveryMethod( n )
	END PROPERTY

	PROPERTY LET User( s )
	END PROPERTY
	
	PROPERTY LET PW( s )
	END PROPERTY
	
	PROPERTY LET From( s )
		m_sFROM = s
	END PROPERTY
	
	PROPERTY LET FromName( s )
		m_sFromName = s
	END PROPERTY
	
	PROPERTY LET ReplyTo( s )
	END PROPERTY
	
	SUB AddRecipient( sAddr )
		m_sTO = m_sTO & ";" & sAddr
	END SUB
	
	SUB AddRecipient2( sAddr, sName )
		m_sTO = m_sTO & ";" & sAddr & " (" & sName & ")"
	END SUB
	
	SUB AddRecipientCC( sAddr )
		m_sCC = m_sCC & ";" & sAddr
	END SUB
	
	SUB AddRecipientCC2( sAddr, sName )
		m_sCC = m_sCC & ";" & sAddr & " (" & sName & ")"
	END SUB
	
	SUB AddRecipientBCC( sAddr )
		m_sBCC = m_sBCC & ";" & sAddr
	END SUB
	
	SUB AddRecipientBCC2( sAddr, sName )
		m_sBCC = m_sBCC & ";" & sAddr & " (" & sName & ")"
	END SUB
	
	FUNCTION AddInlineAttachment( sFile )
		DIM	sAlias
		DIM	sName
		DIM	i
		i = INSTRREV( sFile, "\" )
		IF 0 < i THEN
			sName = MID( sFile, i+1 )
		ELSE
			sName = sFile
		END IF
		m_nInline = m_nInline + 1
		sAlias = "inline" & m_nInline & "_" & sName
		m_oSMTP.AttachURL sFile, sAlias
		AddInlineAttachment = sAlias
	END FUNCTION
	
	SUB AddAttachment( sFile )
		addAttachmentList sFile
	END SUB
	
	PROPERTY LET Subject( s )
		m_oSMTP.Subject = s
	END PROPERTY
	
	PROPERTY LET Body( s )
		IF "<html>" = LCASE(LEFT(s,6))  _
				OR  "<!doctype html" = LCASE(LEFT(s,14)) THEN
			m_oSMTP.Body = s
			m_oSMTP.MailFormat = 0
			m_oSMTP.BodyFormat = 0
		ELSE
			m_oSMTP.Body = s
			m_oSMTP.BodyFormat = 1
		END IF
	END PROPERTY
	
	PROPERTY LET Priority( n )
		DIM	p
		SELECT CASE mailPriority( n )
		CASE k_nMailPriorityNone
			p = 1
		CASE k_nMailPriorityLow
			p = 0
		CASE k_nMailPriorityMedium
			p = 1
		CASE k_nMailPriorityHigh
			p = 2
		CASE ELSE
			p = 1
		END SELECT
		m_oSMTP.Importance = p
	END PROPERTY
	
	FUNCTION Send()
		IF m_sFromName <> "" THEN m_sFROM = m_sFROM & " (" & m_sFromName & ")"
		m_oSMTP.From = m_sFROM
		m_oSMTP.To = MID(m_sTO,2)
		IF "" <> m_sCC THEN m_oSMTP.CC = MID(m_sCC,2)
		IF "" <> m_sBCC THEN m_oSMTP.BCC = MID(m_sBCC,2)
		IF 0 < m_nAttachments THEN
			DIM i
			FOR i = 0 TO m_nAttachments-1
				addAttachmentFile m_aAttachments(i)
			NEXT 'i
		END IF
		ON ERROR Resume Next
		m_oSMTP.Send
		IF Err <> 0 THEN
			Send = Err.Number & ":" & Err.Description
		ELSE
			Send = "0:Success"
		END IF
		ON ERROR Goto 0
		m_sTO = ""
		m_sCC = ""
		m_sBCC = ""
	END FUNCTION

END CLASS



	CONST k_cdoSendUsingPickup = 1
	CONST k_cdoSendUsingPort = 2
	CONST k_cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
	CONST k_cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
	CONST k_cdoSMTPconnectiontimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
	CONST k_cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
	CONST k_cdoSMTPServerPickupDirectory = "http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory"
	CONST k_cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
	
	CONST k_cdoAnonymous = 0
	CONST k_cdoBasic = 1
	CONST k_cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
	CONST k_cdoSMTPAccountName = "http://schemas.microsoft.com/cdo/configuration/smtpaccountname"
	CONST k_cdoSMTPEmailAddress = "http://schemas.microsoft.com/cdo/configuration/smtpemailaddress"
	CONST k_cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
	CONST k_cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
	

CLASS CDOSYSSMTP

	PRIVATE	m_oSMTP
	PRIVATE	m_sServer
	PRIVATE m_nDeliveryMethod
	PRIVATE m_sUser
	PRIVATE m_sPW
	
	PRIVATE	m_sFROM
	PRIVATE m_sFromName
	PRIVATE	m_sTO
	PRIVATE	m_sCC
	PRIVATE	m_sBCC
	
	PRIVATE	m_nInline
	PRIVATE m_aAttachments()
	PRIVATE m_nAttachments
	
	SUB Class_Initialize()
		SET m_oSMTP = Nothing
		m_sServer = ""
		m_nDeliveryMethod = k_nMailDeliveryDefault
		m_sUser = ""
		m_sPW = ""
		m_sFROM = ""
		m_sFromName = ""
		m_sTO = ""
		m_sCC = ""
		m_sBCC = ""
		m_nInline = 0
		REDIM m_aAttachments(5)
		m_nAttachments = 0
		'Response.Write "Class_Initialize: CDOSYS<br>"
		'ON ERROR Resume Next
		'SET m_oSMTP = Server.CreateObject("CDO.Message")
		'ON ERROR Goto 0
	END SUB
	
	SUB Class_Terminate()
		'Response.Write "Class_Terminate: CDOSYS<br>"
		SET m_oSMTP = Nothing
	END SUB
	
	PRIVATE SUB addAttachmentList( s )
		IF UBOUND(m_aAttachments) <= m_nAttachments+1 THEN
			REDIM PRESERVE m_aAttachments(UBOUND(m_aAttachments)+5)
		END IF
		m_aAttachments(m_nAttachments) = s
		m_nAttachments = m_nAttachments + 1
	END SUB
	
	PUBLIC PROPERTY GET SMTP
		'IF Nothing IS m_oSMTP THEN Response.Write "Problem creating component CDO.Message<br>"
		SET SMTP = m_oSMTP
	END PROPERTY

	PROPERTY SET SMTP( o )
		SET m_oSMTP = o
	END PROPERTY

	PROPERTY LET Server( s )
		m_sServer = s
	END PROPERTY

	PROPERTY LET DeliveryMethod( n )
		m_nDeliveryMethod = n
	END PROPERTY

	PROPERTY LET User( s )
		m_sUser = s
	END PROPERTY
	
	PROPERTY LET PW( s )
		m_sPW = s
	END PROPERTY
	
	PROPERTY LET From( s )
		m_sFROM = s
	END PROPERTY
	
	PROPERTY LET FromName( s )
		m_sFromName = s
	END PROPERTY
	
	PROPERTY LET ReplyTo( s )
		WITH m_oSMTP.Fields
			.item("urn:schemas:mailheader:reply-to") = s
			.item("urn:schemas:httpmail:reply-to") = s
			.Update
		END WITH
	END PROPERTY
	
	SUB AddRecipient( sAddr )
		m_sTO = m_sTO & "," & sAddr
	END SUB
	
	SUB AddRecipient2( sAddr, sName )
		m_sTO = m_sTO & ",""" & sName & """ <" & sAddr & ">"
	END SUB
	
	SUB AddRecipientCC( sAddr )
		m_sCC = m_sCC & "," & sAddr
	END SUB
	
	SUB AddRecipientCC2( sAddr, sName )
		m_sCC = m_sCC & ",""" & sName & """ <" & sAddr & ">"
	END SUB
	
	SUB AddRecipientBCC( sAddr )
		m_sBCC = m_sBCC & "," & sAddr
	END SUB
	
	SUB AddRecipientBCC2( sAddr, sName )
		m_sBCC = m_sBCC & ",""" & sName & """ <" & sAddr & ">"
	END SUB
	
	FUNCTION AddInlineAttachment( sFile )
		DIM	sAlias
		DIM	sName
		DIM	i
		i = INSTRREV( sFile, "\" )
		IF 0 < i THEN
			sName = MID( sFile, i+1 )
		ELSE
			sName = sFile
		END IF
		m_nInline = m_nInline + 1
		sAlias = "inline" & m_nInline & "_" & sName
		'sAlias = sName
		DIM	e
		'ON ERROR Resume Next
		'response.write sFile & "<br>"
		'response.flush
		m_oSMTP.AddRelatedBodyPart sFile, sAlias, 0
		e = Err
		'ON ERROR GOTO 0
		IF e = 0 THEN
			AddInlineAttachment = sAlias
		ELSE
			AddInlineAttachment = ""
		END IF
	END FUNCTION
	
	SUB AddAttachment( sFile )
		' we queue up the list of attachments because they need to be made
		' after the body and any inline attachments have been made.
		addAttachmentList sFile
	END SUB
	
	PROPERTY LET Subject( s )
		m_oSMTP.Subject = s
	END PROPERTY
	
	PROPERTY LET Body( s )
		m_oSMTP.AutoGenerateTextBody = FALSE
		IF "<html>" = LCASE(LEFT(s,6))  _
				OR  "<!doctype html" = LCASE(LEFT(s,14)) THEN
			m_oSMTP.HtmlBody = s
		ELSE
			m_oSMTP.TextBody = s
		END IF
	END PROPERTY
	
	PROPERTY LET Priority( n )
		DIM	p
		DIM	i
		SELECT CASE mailPriority( n )
		CASE k_nMailPriorityNone
			i = 1
			p = 0
		CASE k_nMailPriorityLow
			i = 0
			p = -1
		CASE k_nMailPriorityMedium
			i = 1
			p = 0
		CASE k_nMailPriorityHigh
			i = 2
			p = 1
		CASE ELSE
			i = 1
			p = 0
		END SELECT
		WITH m_oSMTP.Fields
			.item("urn:schemas:httpmail:importance") = i
			.item("urn:schemas:httpmail:priority") = p
			.Update
		END WITH
	END PROPERTY
	
	SUB wd( p, v )
		'Response.Write p & ": " & v & "<br>" & vbCRLF
	END SUB
	
	FUNCTION Send()
		DIM oConfig
		SET oConfig = CreateObject("CDO.Configuration")
		with oConfig.Fields
			wd "m_sServer", m_sServer
			IF m_nDeliveryMethod = k_nMailDeliveryPickup THEN
				.Item(k_cdoSendUsingMethod) = k_cdoSendUsingPickup
				.Item(k_cdoSMTPServer) = "localhost"
				'.Item(k_cdoSMTPServerPickupDirectory) = Request.ServerVariables("APPL_MD_PATH")
			ELSEIF m_nDeliveryMethod = k_nMailDeliveryPort THEN
				.Item(k_cdoSendUsingMethod) = k_cdoSendUsingPort
				.Item(k_cdoSMTPServer) = m_sServer
				.Item(k_cdoSMTPconnectiontimeout) = 20
				.Item(k_cdoSMTPServerPort) = 25
			END IF
			IF "" <> m_sUser  AND  "" <> m_sPW THEN
				.Item(k_cdoSMTPAuthenticate) = k_cdoBasic
				.Item(k_cdoSMTPAccountName) = m_sUser
				.Item(k_cdoSendUserName) = m_sUser
				.Item(k_cdoSendPassword) = m_sPW
				.Item(k_cdoSMTPUseSSL) = FALSE
			ELSE
				.Item(k_cdoSMTPAuthenticate) = k_cdoAnonymous
			END IF
			.Update()
		end with
		SET m_oSMTP.Configuration = oConfig
		SET oConfig = Nothing
		IF m_sFromName <> "" THEN m_sFROM = """" & m_sFromName & """ <" & m_sFrom & ">"
		wd "m_sFROM", m_sFROM
		m_oSMTP.From = m_sFROM
		wd "m_sTO", MID(m_sTO,2)
		m_oSMTP.To = MID(m_sTO,2)
		IF "" <> m_sCC THEN
			wd "m_sCC", MID(m_sCC,2)
			m_oSMTP.CC = MID(m_sCC,2)
		END IF
		IF "" <> m_sBCC THEN
			wd "m_sBCC", MID(m_sBCC,2)
			m_oSMTP.BCC = MID(m_sBCC,2)
		END IF
		
		IF 0 < m_nAttachments THEN
			m_oSMTP.BodyPart.ContentMediaType = "multipart/mixed"
			m_oSMTP.MimeFormatted = TRUE

			DIM i
			ON ERROR Resume Next
			FOR i = 0 TO m_nAttachments-1
				m_oSMTP.AddAttachment m_aAttachments(i)
			NEXT 'i
			ON ERROR Goto 0
		END IF
		ON ERROR Resume Next
		'	Response.Write "<pre>" & vbCRLF
		'	Response.Write m_oSMTP.GetStream.readtext
		'	Response.Write "</pre>" & vbCRLF
		m_oSMTP.Send
		IF Err <> 0 THEN
			Send = Err.Number & ":" & Err.Description
		ELSE
			Send = "0:Success"
		END IF
		ON ERROR Goto 0
		m_sTO = ""
		m_sCC = ""
		m_sBCC = ""
	END FUNCTION

END CLASS




CLASS ASPMailSMTP

	PRIVATE	m_oSMTP
	
	SUB Class_Initialize()
		SET m_oSMTP = Nothing
		'Response.Write "Class_Initialize: ASPMail<br>"
		'ON ERROR Resume Next
		'SET m_oSMTP = Server.CreateObject("SMTPsvg.Mailer")
		'ON ERROR Goto 0
	END SUB
	
	SUB Class_Terminate()
		'Response.Write "Class_Terminate: ASPMail<br>"
		SET m_oSMTP = Nothing
	END SUB
	
	PROPERTY GET SMTP
		'IF Nothing IS m_oSMTP THEN Response.Write "Problem creating component SMTPsvg.Mailer<br>"
		SET SMTP = m_oSMTP
	END PROPERTY
	
	PROPERTY SET SMTP( o )
		SET m_oSMTP = o
	END PROPERTY
	
	PROPERTY LET Server( s )
		m_oSMTP.RemoteHost = s
	END PROPERTY

	PROPERTY LET DeliveryMethod( n )
	END PROPERTY

	PROPERTY LET User( s )
	END PROPERTY
	
	PROPERTY LET PW( s )
	END PROPERTY
	
	PROPERTY LET From( s )
		m_oSMTP.FromAddress = s
	END PROPERTY
	
	PROPERTY LET FromName( s )
		m_oSMTP.FromName = s
	END PROPERTY
	
	PROPERTY LET ReplyTo( s )
		m_oSMTP.ReplyTo = s
	END PROPERTY
	
	SUB AddRecipient( sAddr )
		m_oSMTP.AddRecipient sAddr, sAddr
	END SUB
	
	SUB AddRecipient2( sAddr, sName )
		m_oSMTP.AddRecipient sName, sAddr
	END SUB
	
	SUB AddRecipientCC( sAddr )
		m_oSMTP.AddCC sAddr, sAddr
	END SUB
	
	SUB AddRecipientCC2( sAddr, sName )
		m_oSMTP.AddCC sName, sAddr
	END SUB
	
	SUB AddRecipientBCC( sAddr )
		m_oSMTP.AddBCC sAddr, sAddr
	END SUB
	
	SUB AddRecipientBCC2( sAddr, sName )
		m_oSMTP.AddBCC sName, sAddr
	END SUB

	FUNCTION AddInlineAttachment( sFile )
		m_oSMTP.AddAttachment sFile
		AddInlineAttachment = sFile
	END FUNCTION
	
	SUB AddAttachment( sFile )
		m_oSMTP.AddAttachment sFile
	END SUB
	

	PROPERTY LET Subject( s )
		m_oSMTP.Subject = s
	END PROPERTY
	
	PROPERTY LET Body( s )
		m_oSMTP.BodyText = s
		IF "<html>" = LCASE(LEFT(s,6))  _
				OR  "<!doctype html" = LCASE(LEFT(s,14)) THEN
			m_oSMTP.ContentType = "text/html"
		ELSE
			m_oSMTP.ContentType = "text/plain"
		END IF
	END PROPERTY

	PROPERTY LET Priority( n )
		DIM	p
		SELECT CASE mailPriority( n )
		CASE k_nMailPriorityNone
			p = 3
		CASE k_nMailPriorityLow
			p = 5
		CASE k_nMailPriorityMedium
			p = 3
		CASE k_nMailPriorityHigh
			p = 1
		CASE ELSE
			p = 3
		END SELECT
		m_oSMTP.Priority = p
	END PROPERTY

	FUNCTION Send()
		IF m_oSMTP.SendMail() THEN
			Send = "0:Success"
		ELSE
			Send = "-1:" & m_oSMTP.Response
		END IF
	END FUNCTION

END CLASS



SUB mailLoadNames( oSMTP, sFrom, sReplyTo, sTo, sCC, sBCC )
	DIM	i
	DIM	sAddr
	DIM	sName
	DIM	sFromName
	DIM	sMailServer
	DIM	aToSplit
	

	parseAddr sFromName, sAddr, sFrom

	
	'Response.Write "From: " & sAddr
	oSMTP.From = sAddr
	IF 0 < LEN(sFromName) THEN
		'Response.Write " -- " & sFromName
		oSMTP.FromName = sFromName
	END IF
	'Response.Write "<br>" & vbCRLF
			
	IF 0 < LEN(sReplyTo) THEN
		'Response.Write "ReplyTo: " & sReplyTo & "<br>" & vbCRLF
		parseAddr sName, sAddr, sReplyTo
		oSMTP.ReplyTo = sAddr
	END IF
	
	aToSplit = SPLIT(sTo, ";", -1, vbTextCompare)
	FOR i = LBOUND(aToSplit) TO UBOUND(aToSplit)
		parseAddr sName, sAddr, TRIM(aToSplit(i))
		IF 0 < LEN(sAddr) THEN
			IF 0 < LEN(sName) THEN
				oSMTP.AddRecipient2 sAddr, sName
			ELSE
				oSMTP.AddRecipient sAddr
			END IF
		END IF
	NEXT 'i
			
	IF 0 < LEN(sCC) THEN
		aToSplit = SPLIT(sCC, ";", -1, vbTextCompare)
		FOR i = LBOUND(aToSplit) TO UBOUND(aToSplit)
			parseAddr sName, sAddr, TRIM(aToSplit(i))
			IF 0 < LEN(sAddr) THEN
				IF 0 < LEN(sName) THEN
					oSMTP.AddRecipientCC2 sAddr, sName
				ELSE
					oSMTP.AddRecipientCC sAddr
				END IF
			END IF
		NEXT 'i
	END IF
			
	IF 0 < LEN(sBCC) THEN
		aToSplit = SPLIT(sBCC, ";", -1, vbTextCompare)
		FOR i = LBOUND(aToSplit) TO UBOUND(aToSplit)
			parseAddr sName, sAddr, TRIM(aToSplit(i))
			IF 0 < LEN(sAddr) THEN
				IF 0 < LEN(sName) THEN
					oSMTP.AddRecipientBCC2 sAddr, sName
				ELSE
					oSMTP.AddRecipientBCC sAddr
				END IF
			END IF
		NEXT 'i
	END IF
			
END SUB



FUNCTION mailsender( oSMTP, sFrom, sReplyTo, sTo, sCC, sBCC, sSubject, ByRef sText )
	
	oSMTP.Server = mailServer()

	mailLoadNames oSMTP, sFrom, sReplyTo, sTo, sCC, sBCC
			
	'Response.Write "Subject: " & Server.HTMLEncode(sSubject) & "<br>" & vbCRLF
	oSMTP.Subject = sSubject
			
	oSMTP.Body = sText
			
	'send it
	DIM	sError
	sError = oSMTP.Send
	IF "0:" = LEFT(sError,2) THEN
		mailsender = 0
	ELSE
		mailsender = -1
		Response.Write "<p>" & Server.HTMLEncode(sError) & "</p>" & vbCRLF
	END IF
END FUNCTION


FUNCTION mailMakeSMTP()

	DIM oSMTP
	DIM	o
	
	ON ERROR Resume Next
	SET oSMTP = Nothing
	SET o = Nothing
	IF Nothing IS o THEN
		SET o = Server.CreateObject("CDO.Message")
		IF NOT Nothing IS o THEN
			SET oSMTP = new CDOSYSSMTP
			SET oSMTP.SMTP = o
		END IF
	END IF
	IF Nothing IS o THEN
		SET o = Server.CreateObject("CDONTS.NewMail")
		IF NOT Nothing IS o THEN
			SET oSMTP = new CDONTSSMTP
			SET oSMTP.SMTP = o
		END IF
	END IF
	
	IF Nothing IS o THEN
		SET o = Server.CreateObject("SMTPsvg.Mailer")
		IF NOT Nothing IS o THEN
			SET oSMTP = new ASPMailSMTP
			SET oSMTP.SMTP = o
		END IF
	END IF
	IF Nothing IS o THEN
		SET o = Server.CreateObject("Persits.MailSender")
		IF NOT Nothing IS o THEN
			SET oSMTP = new PersistsMailSMTP
			SET oSMTP.SMTP = o
		END IF
	END IF
	IF Nothing IS o THEN
		SET o = Server.CreateObject("JMail.Message")
		IF NOT Nothing IS o THEN
			SET oSMTP = new JMailSMTP
			SET oSMTP.SMTP = o
		END IF
	END IF
	ON ERROR GOTO 0
	IF Nothing IS oSMTP  OR  Nothing IS oSMTP.SMTP THEN
		SET mailMakeSMTP = Nothing
	ELSE
		SET mailMakeSMTP = oSMTP
	END IF
	SET oSMTP = Nothing
	SET o = Nothing

END FUNCTION


FUNCTION mailsend( sFrom, sReplyTo, sTo, sCC, sBCC, sSubject, ByRef sText )

	DIM oSMTP
	
	SET oSMTP = mailMakeSMTP()
	IF Nothing IS oSMTP THEN
		mailsend = -1
	ELSE
		mailsend = mailsender( oSMTP, sFrom, sReplyTo, sTo, sCC, sBCC, sSubject, sText )
		'IF 0 = mailsend THEN
		'	Response.Write "<p>Successfully sent using: " & TypeName(oSMTP.SMTP) & "</p>" & vbCRLF
		'END IF
	END IF
	SET oSMTP = Nothing

END FUNCTION


%>