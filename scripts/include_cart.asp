<%




FUNCTION cartSessionKey()
	DIM	i
	DIM	sSessionKey
	sSessionKey = "CARTUser" & Request.ServerVariables("PATH_INFO")
	i = INSTRREV(sSessionKey,"/")
	IF 0 < i THEN sSessionKey = LEFT(sSessionKey,i-1)
	cartSessionKey = sSessionKey
END FUNCTION


FUNCTION makeGUID()
	makeGUID = ""
	DIM	oTypeLib
	
	SET oTypeLib = Server.CreateObject("Scriptlet.Typelib")
	IF NOT Nothing IS oTypeLib THEN
		DIM	sGUID
		
		sGUID = TRIM(oTypeLib.Guid)
		sGUID = MID(LEFT(sGUID,INSTR(sGUID,"}")-1),2)
		SET oTypeLib = Nothing
		makeGUID = TRIM(sGUID)
	END IF
END FUNCTION

DIM	sCartCookie

FUNCTION doCartCookie()
	DIM	sCookie
	DIM	sSessionKey
	sSessionKey = cartSessionKey()
	
	sCookie = TRIM(Session.Contents(sSessionKey))
	IF "" = sCookie THEN
		sCookie = TRIM(Request.Cookies(sSessionKey))
		IF "" = sCookie THEN
			sCookie = makeGUID()
		END IF
	END IF
	Response.Cookies(sSessionKey) = sCookie
	Response.Cookies(sSessionKey).Expires = DateAdd( "d", 7*8, NOW )
	Session(sSessionKey) = sCookie
	
	doCartCookie = sCookie

END FUNCTION
sCartCookie = doCartCookie()

'Response.Write "<p>" & sCartCookie & "</p>" & vbCRLF
'Response.Flush



%>