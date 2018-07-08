<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon_defs.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<%


SUB recordGoodLogon()
	' record logon
	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"* " _
		&	"FROM " _
		&		dbQ("authorizelog") & " " _
		&	"WHERE " _
		&		dbQ("UserID") & " = " & g_auth_nUserID & " " _
		&	";"
	DIM	oRS
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		IF oRS.EOF THEN		' no record found
			oRS.AddNew
			oRS.Fields("UserID").Value = fieldNumber(g_auth_nUserID)
		END IF
		oRS.Fields("DateLogon").Value = NOW
		oRS.Fields("DateFail").Value = fieldDate(CDATE("12:00:00 AM"))
		oRS.Fields("FailCount").Value = 0
		oRS.Update
		oRS.Close
	END IF
	SET oRS = Nothing
END SUB


FUNCTION getFlagsList()
	DIM	s
	s = ""
	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"authorizeresource.Abbrev " _
		&	"FROM " _
		&		"(authorizeresource " _
		&		"INNER JOIN authorizeresourcemap " _
		&			"ON authorizeresourcemap.ResourceID = authorizeresource.RID) " _
		&			"INNER JOIN authorizegroupmap " _
		&				"ON authorizegroupmap.GroupID = authorizeresourcemap.GroupID " _
		&	"WHERE " _
		&		"authorizegroupmap.UserID = " & g_auth_nUserID & " " _
		&	";"
	DIM	oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		DIM	oAbv
		SET oAbv = oRS.Fields("Abbrev")
		DIM	sAbv
		oRS.MoveFirst
		DO UNTIL oRS.EOF
			sAbv = recString(oAbv)
			IF "" <> sAbv THEN s = s & "," & sAbv
			oRS.MoveNext
		LOOP
		oRS.Close
		SET oRS = Nothing
		s = MID(s, 2)	' get rid of leading comma
	END IF
	getFlagsList = s
END FUNCTION




DIM	nError
DIM	sError
nError = 0
sError = ""


DIM	sUsername
DIM	sPassword
DIM	sRemember

sUsername = TRIM(Request("Username"))
sPassword = TRIM(Request("Password"))
sRemember = Request("Remember")



g_auth_bAccept = FALSE

DIM	sSelect
DIM	oRS

DIM	nMxsUsername
DIM	nMxsPassword

nMxsUsername = dbGetFieldSize( g_DC, "authorizeusers", "UserName" )
nMxsPassword = dbGetFieldSize( g_DC, "authorizeusers", "Password" )



IF LEN(sUsername) < 3  OR  nMxsUsername < LEN(sUsername) THEN
	nError = kAuth_Username_Length
ELSEIF 0 < INSTR(sUsername,"'")  OR  0 < INSTR(sUsername,"--")  OR  0 < INSTR(sUsername,"""") THEN
	nError = kAuth_Username_BadChar
ELSEIF LEN(sPassword) < 6  OR  nMxsPassword < LEN(sPassword) THEN
	nError = kAuth_Password_Length
ELSEIF 0 < INSTR(sPassword,"'")  OR  0 < INSTR(sPassword,"--")  OR  0 < INSTR(sPassword,"""") THEN
	nError = kAuth_Password_BadChar
END IF

IF 0 = nError THEN



sSelect = "" _
	&	"SELECT " _
	&		"* " _
	&	"FROM " _
	&		dbQ("authorizeusers") & " " _
	&	"WHERE " _
	&		dbQ("Username") & " = '" & sUsername & "' " _
	&		"AND " & dbQ("Password") & " = '" & sPassword & "' " _
	&	";"

SET oRS = dbQueryReadEx( g_DC, sSelect )
IF NOT oRS IS Nothing THEN
	oRS.MoveFirst
	IF NOT oRS.EOF THEN
		DIM	nRID
		g_auth_nUserID = recNumber( oRS.Fields("RID") )
		g_auth_sUsername = recString( oRS.Fields("Username") )
		g_auth_sPassword = recString( oRS.Fields("Password") )
		g_auth_bAccept = TRUE
		g_auth_sFlags = ""	' future
		IF "" = sRemember THEN
			g_auth_bRemember = FALSE
		ELSE
			SELECT CASE LCASE(LEFT(sRemember,1))
			CASE "y", "t", "1"
				g_auth_bRemember = TRUE
			CASE ELSE
				g_auth_bRemember = FALSE
			END SELECT
		END IF
	END IF
	oRS.Close
END IF
SET oRS = Nothing



' record logon
IF g_auth_bAccept THEN recordGoodLogon



' get the list of allowed resources
IF g_auth_bAccept THEN
	g_auth_sFlags = getFlagsList()
	authSetAllCookies
END IF

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

IF g_auth_bAccept THEN
	Response.Redirect "./"
ELSE
	Response.Redirect "logon.asp?error=" & kAuth_UnknownUser
END IF


ELSE
	Response.Redirect "logon.asp?error=" & nError
END IF


%>
