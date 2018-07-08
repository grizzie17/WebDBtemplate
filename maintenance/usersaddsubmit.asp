<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon_defs.asp"-->
<%

SUB ErrorReturn( nError )
	DIM	s
	s = "" _
		&	"&f=" & Server.URLEncode(sFirstname) _
		&	"&l=" & Server.URLEncode(sLastname) _
		&	"&u=" & Server.URLEncode(sUsername) _
		&	"&m=" & Server.URLEncode(sEmail) _
		&	"&d=" & Server.URLEncode(sDescription)

	Response.Redirect "usersadd.asp?error=" & nError & s
END SUB






DIM	sTabID
DIM	sFirstname
DIM	sLastname
DIM	sUsername
DIM	sPassword
DIM	sEmail
DIM	sDescription

sFirstname = Request("Firstname")
sLastname = Request("Lastname")
sUsername = Request("Username")
sPassword = Request("Password")
sEmail = Request("Email")
sDescription = Request("Description")



IF LEN(sUsername) < 3 THEN
	ErrorReturn kAuth_Username_Length
END IF

IF 0 < INSTR(sUsername,"'")  OR 0 < INSTR(sUsername,"""")  OR 0 < INSTR(sPassword,"--") THEN
	ErrorReturn kAuth_Username_BadChar
END IF

IF LEN(sPassword) < 6 THEN
	ErrorReturn kAuth_Password_Length
END IF

IF 0 < INSTR(sPassword,"'")  OR 0 < INSTR(sPassword,"""")  OR 0 < INSTR(sPassword,"--") THEN
	ErrorReturn kAuth_Password_BadChar
END IF

IF LEN(sEmail) < 6 THEN
	ErrorReturn kAuth_Email_Length
END IF

IF 0 = INSTR(sEmail,"@")  OR  0 < INSTR(sEmail,"'") OR  0 < INSTR(sEmail,"""")  OR  0 < INSTR(sEmail,"--") THEN
	ErrorReturn kAuth_Email_BadChar
END IF



%>
<!--#include file="scripts\include_db_begin.asp"-->
<%

DIM	nError
DIM	sError
nError = 0
sError = ""


DIM	sSelect
DIM	oRS
	
sSelect = "" _
	&	"SELECT " _
	&		"* " _
	&	"FROM " _
	&		"authorizeusers " _
	&	"WHERE " _
	&		"Username = '" & sUsername & "' " _
	&		"OR " _
	&		"Email LIKE '" & sEmail & "' " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )
IF NOT oRS IS Nothing THEN
	IF NOT oRS.EOF THEN
		DO UNTIL oRS.EOF
			IF sUsername = recString(oRS.Fields("Username")) THEN
				nError = kAuth_Username_Exists
				EXIT DO
			ELSEIF LCASE(sEmail) = recString(oRS.Fields("Email")) THEN
				nError = kAuth_Email_Exists
			END IF
			oRS.MoveNext
		LOOP	
	END IF
	oRS.Close
	SET oRS = Nothing
END IF




IF 0 = nError THEN


	sSelect = "" _
		&	"SELECT " _
		&	" * " _
		&	"FROM authorizeusers " _
		&	"WHERE " _
		&	" RID=0" _
		&	";"

	DIM	nUserID

	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.AddNew
		
		oRS.Fields("Firstname").Value = sFirstname
		oRS.Fields("Lastname").Value = sLastname
		oRS.Fields("Username").Value = sUsername
		oRS.Fields("Password").Value = sPassword
		oRS.Fields("Email").Value = sEmail
		oRS.Fields("Description").Value = fieldString(sDescription)
		oRS.Fields("DateCreated").Value = fieldDate(NOW)
		oRS.Fields("DateModified").Value = fieldDate(NOW)
		oRS.Fields("Disabled").Value = 0
		oRS.Update
		nUserID = oRS.Fields("RID").Value

		oRS.Close
		SET oRS = Nothing
	END IF


END IF


%>
<!--#include file="scripts\include_db_end.asp"-->
<%

IF 0 = nError THEN

	Response.Redirect "usersedit.asp?user=" & nUserID

ELSE

	ErrorReturn nError

END IF

%>