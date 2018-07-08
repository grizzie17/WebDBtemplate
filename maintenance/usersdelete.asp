<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<!--#include file="scripts\include_relations.asp"-->
<%


DIM	sID
sID = TRIM(Request("tab"))



IF "" <> sID THEN

	DIM	oRS
	DIM	sSelect
	
	IF 0 < INSTR(sID,",") THEN
		sWhere = " UserID IN (" & sID & ") "
	ELSE
		sWhere = " UserID=" & sID & " "
	END IF
	
	sSelect = "" _
		&	"SELECT " _
		&		"* " _
		&	"FROM " _
		&		"authorizelog " _
		&	"WHERE " _
		&		sWhere & " " _
		&	";"

	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		DO UNTIL oRS.EOF
			oRS.Delete 1
			oRS.MoveNext
		LOOP
		oRS.Close
	END IF
	SET oRS = Nothing


	sSelect = "" _
		&	"SELECT " _
		&		"* " _
		&	"FROM " _
		&		"authorizegroupmap " _
		&	"WHERE " _
		&		sWhere & " " _
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


	DIM	sWhere
	IF 0 < INSTR(sID,",") THEN
		sWhere = " RID IN (" & sID & ") "
	ELSE
		sWhere = " RID=" & sID & " "
	END IF


	sSelect = "SELECT " _
		&		"* " _
		&	"FROM authorizeusers " _
		&	"WHERE " _
		&		sWhere _
		&	";"

	DIM	g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT g_RS.EOF THEN
			
		g_RS.MoveFirst		
		DO UNTIL g_RS.EOF			
			g_RS.Delete 1 'adAffectCurrent
			g_RS.MoveNext
		LOOP

	END IF



	
	Set g_RS.ActiveConnection = Nothing
	g_RS.Close
	SET g_RS = Nothing

END IF

SET g_RS = Nothing

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "users.asp"

%>