<%

CONST	kAuth_Username_Length	= -100
CONST	kAuth_Username_BadChar	= -101
CONST	kAuth_Username_Exists	= -102

CONST	kAuth_Password_Length	= -200
CONST	kAuth_Password_BadChar	= -201

CONST	kAuth_Email_Length		= -300
CONST	kAuth_Email_BadChar		= -301
CONST	kAuth_Email_Exists		= -302

CONST	kAuth_UnknownUser		= -900



FUNCTION authErrorString( n )
	DIM	s
	SELECT CASE n
	CASE kAuth_Username_Length
		s = "Username must be at least 3 characters long"
	CASE kAuth_Username_BadChar
		s = "Username contains a bad character - please restrict to letters and numbers"
	CASE kAuth_Username_Exists
		s = "Username already exists - please choose a different name"
	CASE kAuth_Password_Length
		s = "Password must be at least 6 characters"
	CASE kAuth_Password_BadChar
		s = "Password contains a bad character - please avoid the use of the appostrophe"
	CASE kAuth_Email_Length
		s = "Email must be specified when creating a new user"
	CASE kAuth_Email_BadChar
		s = "Email appears to be badly formed - please correct"
	CASE kAuth_Email_Exists
		s = "Email address already used by another user - please use a different email address"
	CASE kAuth_UnknownUser
		s = "Username or Password is unrecognized - please check your spelling"
	END SELECT
	authErrorString = s
END FUNCTION





DIM	g_auth_sUsername
DIM	g_auth_sPassword
DIM	g_auth_nUserID
DIM	g_auth_bRemember
DIM	g_auth_bAccept
DIM	g_auth_bGood
DIM	g_auth_sFlags




FUNCTION authHasAccess( s )

	IF 0 < INSTR( ","&g_auth_sFlags&",", ","&s&"," ) THEN
		authHasAccess = TRUE
	ELSE
		authHasAccess = FALSE
	END IF

END FUNCTION




FUNCTION authCookie( s )
	authCookie = Request.Cookies( g_sCookiePrefix & "_" & s )
	'Response.Write s & " = " & authCookie & "<br>" & vbCRLF
	'Response.Flush
END FUNCTION

FUNCTION authCookieBool( s )
	DIM	sValue
	sValue = authCookie( s )
	IF "" = sValue THEN
		authCookieBool = FALSE
	ELSE
		SELECT CASE LCASE(LEFT(sValue,1))
		CASE "y", "t", "1"
			authCookieBool = TRUE
		CASE ELSE
			authCookieBool = FALSE
		END SELECT
	END IF
END FUNCTION

FUNCTION authCookieNumber( s )
	DIM	sValue
	sValue = authCookie( s )
	IF "" = sValue THEN
		authCookieNumber = 0
	ELSEIF ISNUMERIC( sValue ) THEN
		authCookieNumber = CLNG(sValue)
	ELSE
		authCookieNumber = 0
	END IF
END FUNCTION

SUB authSetCookie( sCookie, sValue )
	Response.Cookies( g_sCookiePrefix & "_" & sCookie ) = CSTR(sValue)
END SUB

SUB authSetCookieExpire( sCookie, sValue, dExp )
	authSetCookie sCookie, sValue
	Response.Cookies( g_sCookiePrefix & "_" & sCookie ).Expires = dExp
END SUB

SUB authSetAllCookies()
	DIM	dExp
	dExp = DateAdd( "d", 14, NOW )
	
	authSetCookie "Accept", g_auth_bAccept
	authSetCookie "Flags", g_auth_sFlags
	authSetCookie "Username", g_auth_sUsername
	
	authSetCookieExpire "UserID", g_auth_nUserID, dExp
	authSetCookieExpire "PW", g_auth_sPassword, dExp
	authSetCookieExpire "Remember", g_auth_bRemember, dExp
END SUB



%>