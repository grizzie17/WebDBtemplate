<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<%


DIM	sBrandID
sBrandID = TRIM(Request("brand"))

IF "" <> sBrandID THEN

	DIM	sBrandWhere
	IF 0 < INSTR(sBrandID,",") THEN
		sBrandWhere = " BrandID IN (" & sBrandID & ") "
	ELSE
		sBrandWhere = " BrandID=" & sBrandID & " "
	END IF


	DIM	sSelect
	sSelect = "SELECT " _
		&		"* " _
		&	"FROM Brands " _
		&	"WHERE " _
		&		sBrandWhere _
		&	";"

	DIM	g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT g_RS.EOF THEN
	
		DIM sLogo
		DIM	oLogo
		SET oLogo = g_RS.Fields("Logo")

		picture_findPath "Brands"
		
		g_RS.MoveFirst		
		DO UNTIL g_RS.EOF
			sLogo = recString(oLogo)
			IF "" <> sLogo THEN
	
				picture_delete sLogo
	
			END IF
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

Response.Redirect "brands.asp"

%>