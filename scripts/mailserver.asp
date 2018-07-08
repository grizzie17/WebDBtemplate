<%
FUNCTION mailServer()
	DIM	sServer
	DIM	sAddr
	DIM	i
	mailServer = ""
	sServer = Request.ServerVariables("SERVER_NAME")
	IF "localhost" = LCASE(sServer) then
		mailServer = "localhost"
	ELSE
		ON ERROR Resume Next
		IF NOT ISEMPTY(g_sMailServer) THEN
			mailServer = g_sMailServer
		END IF
		ON ERROR Goto 0
		IF "" = mailServer THEN
			'sAddr = Request.ServerVariables("LOCAL_ADDR")
			'IF "" <> sAddr THEN
			'	mailServer = sAddr
			'ELSE
				DIM a
				a = SPLIT(sServer,".")
				IF 0 < UBOUND(a) THEN
					mailServer = "mail"
					FOR i = UBOUND(a)-1 TO UBOUND(a)
						mailServer = mailServer & "." & a(i)
					NEXT 'i
				ELSE
					mailServer = "mail." & sServer
				END IF
			'END IF
		END IF
	END IF
	'Response.Write "<p>mailServer = " & mailServer & "</p>" & vbCRLF
END FUNCTION

%>