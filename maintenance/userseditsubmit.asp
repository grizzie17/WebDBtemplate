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

	Response.Redirect "usersedit.asp?error=" & nError & s
END SUB






DIM	sUserID
DIM	sFirstname
DIM	sLastname
DIM	sUsername
DIM	sPassword
DIM	sEmail
DIM	sDescription
DIM	sDisabled

sUserID = Request("UserID")
sFirstname = Request("Firstname")
sLastname = Request("Lastname")
sUsername = Request("Username")
sPassword = Request("Password")
sEmail = Request("Email")
sDescription = Request("Description")
sDisabled = Request("Disabled")


IF LEN(sUserID) < 1  OR  NOT ISNUMERIC(sUserID) THEN
	ErrorReturn kAuth_UnknownUser
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

IF 0 = INSTR(sEmail,"@")  OR  0 < INSTR(sEmail,"'") OR 0 < INSTR(sEmail,"""")  OR  0 < INSTR(sEmail,"--") THEN
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
	&		"RID = " & sUserID & " " _
	&	";"

SET oRS = dbQueryUpdate( g_DC, sSelect )
IF NOT oRS IS Nothing THEN
	IF NOT oRS.EOF THEN
		
		oRS.Fields("Firstname").Value = sFirstname
		oRS.Fields("Lastname").Value = sLastname
		oRS.Fields("Password").Value = sPassword
		oRS.Fields("Email").Value = sEmail
		oRS.Fields("Description").Value = fieldString(sDescription)
		oRS.Fields("DateModified").Value = fieldDate(NOW)
		oRS.Fields("Disabled").Value = fieldBool(sDisabled)
		oRS.Update

		oRS.Close
		SET oRS = Nothing
	ELSE
		nError = 1
	END IF

END IF

' setup relationships with authorization groups
IF 0 = nError THEN

	' first delete all existing records
	sSelect = "" _
		&	"SELECT " _
		&		"* " _
		&	"FROM " _
		&		"authorizegroupmap " _
		&	"WHERE " _
		&		"UserID = " & sUserID & " " _
		&	";"
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		DO UNTIL oRS.EOF
			'Response.Write "delete<br>"
			'Response.Flush
			oRS.Delete 1
			oRS.MoveNext
		LOOP
		oRS.Close
		SET oRS = Nothing
	END IF
	
	
	DIM	nGroupCount
	nGroupCount = 0
	sSelect = "" _
		&	"SELECT " _
		&		"COUNT(*) AS NumGroups " _
		&	"FROM " _
		&		"authorizegroup " _
		&	";"
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		IF NOT oRS.EOF THEN
			nGroupCount = recNumber(oRS.Fields("NumGroups"))
			'Response.Write "Group Count = " & nGroupCount & "<br>"
			'Response.Flush
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
	
	
	sSelect = "" _
		&	"SELECT " _
		&		"* " _
		&	"FROM " _
		&		"authorizegroupmap " _
		&	"WHERE " _
		&		"RID = 0 " _
		&	";"
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		DIM	i
		DIM	sLabel
		DIM	sValue
		FOR i = 1 TO nGroupCount
			sLabel = "Group-" & i
			sValue = Request(sLabel)
			IF "" <> sValue THEN
				IF ISNUMERIC(sValue) THEN
					'Response.Write "add " & sValue & "<br>"
					'Response.Flush
					oRS.AddNew
					oRS.Fields("UserID") = fieldNumber(sUserID)
					oRS.Fields("GroupID") = fieldNumber(sValue)
					oRS.Update
				END IF
			END IF
		NEXT
		oRS.Close
		SET oRS = Nothing
	END IF



END IF


%>
<!--#include file="scripts\include_db_end.asp"-->
<%

IF 0 = nError THEN

	Response.Redirect "users.asp"

ELSE

	ErrorReturn nError

END IF

%>