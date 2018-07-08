<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<!--#include file="scripts\include_relations.asp"-->
<%


DIM	sID
sID = TRIM(Request("category"))

IF "" <> sID THEN

	DIM	sWhere
	IF 0 < INSTR(sID,",") THEN
		sWhere = " CategoryID IN (" & sID & ") "
	ELSE
		sWhere = " CategoryID=" & sID & " "
	END IF


	DIM	sSelect
	sSelect = "SELECT " _
		&		"* " _
		&	"FROM categories " _
		&	"WHERE " _
		&		sWhere _
		&	";"

	DIM	g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT g_RS.EOF THEN
	
		DIM sLogo
		DIM	oLogo
		SET oLogo = g_RS.Fields("Icon")
		
		DIM	oID
		SET oID = g_RS.Fields("CategoryID")

		picture_findPath "Categories"
		
		g_RS.MoveFirst		
		DO UNTIL g_RS.EOF
			sLogo = recString(oLogo)
			IF "" <> sLogo THEN
	
				picture_delete sLogo
	
			END IF
			relations_deleteTabs oID.Value
			
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

Response.Redirect "categories.asp"

%>