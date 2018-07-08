<!--#include file="mailserver.asp"-->
<!--#include file="mailsend.asp"-->
<%


FUNCTION logPage_isUser( sAgent )

	logPage_isUser = TRUE
	IF "" <> sAgent THEN
	
		DIM	oReg
		DIM	bReg

		DIM	sAgentLC
		sAgentLC = LCASE(sAgent)

		SET oReg = NEW RegExp
		oReg.Pattern = "(archive|appie|DoCoMo|bot|robo|crawl|Creep|Fetcher|find|grub|links|Reaper|scooter|Search|seek|sitecheck|slurp|Spider|Synapse|teoma|jeeves|Zeus|zyborg)"
		oReg.IgnoreCase = TRUE
		oReg.Global = TRUE
		bReg = oReg.Test( sAgentLC )
		IF bReg THEN
			logPage_isUser = FALSE
		ELSE
			logPage_isUser = TRUE
		END IF
		
		SET oReg = Nothing
	END IF

END FUNCTION


SUB logPage()
	IF Response.Buffer THEN Response.Flush
	
	DIM	sHost
	sHost = Request.ServerVariables("HTTP_HOST")
	
	IF "localhost" <> LCASE(sHost) THEN
	
		DIM	sAgent
		sAgent = Request.ServerVariables("HTTP_USER_AGENT")
		
		IF logPage_isUser( sAgent ) THEN

			DIM oSMTP
		
			SET oSMTP = mailMakeSMTP()
			IF NOT Nothing IS oSMTP THEN
				oSMTP.Server = "mail.rocketcitywings.org"
				'oSMTP.Server = mailServer()
				
				DIM	sAccount
				DIM	sPW
				
				sAccount = "notifications@rocketcitywings.org"
				sPW = "bear1701"
				
				IF "" <> sAccount  AND  "" <> sPW THEN
					oSMTP.User = sAccount
					oSMTP.PW = sPW
				END IF
				
				DIM	sURL
				sURL = Request.ServerVariables("URL")
				DIM	sQuery
				sQuery = Request.ServerVariables("QUERY_STRING")
				DIM sCookie
				sCookie = Request.ServerVariables("HTTP_COOKIE")
				
				DIM	sPage
				DIM	i
				
				i = INSTRREV( sURL, "/" )
				IF 0 < i THEN
					sPage = MID( sURL, i+1 )
				ELSE
					sPage = sURL
				END IF
	
				mailLoadNames oSMTP, "notifications@rocketcitywings.org", "", "John@GrizzlyWeb.com", "", ""
				
				oSMTP.Subject = "Visitor to RocketCityWings.org" & " - " & sPage
				
				oSMTP.Body = "" _
						&	"Referer = " & Request.ServerVariables("HTTP_REFERER") & vbCRLF _
						&	"Host = " & sHost & vbCRLF _
						&	"URL = " & sURL & vbCRLF _
						&	"Query = " & sQuery & vbCRLF _
						&	"Cookie = " & sCookie & vbCRLF _
						&	sAgent & vbCRLF
				
				
				'send it
				DIM	sError
				sError = oSMTP.Send
			END IF
	
			SET oSMTP = Nothing
		
		END IF

	END IF
END SUB

'logPage


%>
