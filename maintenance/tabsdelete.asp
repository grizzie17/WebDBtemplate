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
<!--#include file="scripts\include_cache.asp"-->
<%


DIM	sID
sID = TRIM(Request("tab"))

IF "" <> sID THEN

	DIM	sWhere
	IF 0 < INSTR(sID,",") THEN
		sWhere = " tabs.TabID IN (" & sID & ") "
	ELSE
		sWhere = " tabs.TabID=" & sID & " "
	END IF


	DIM	sSelect
	sSelect = "SELECT " _
		&		"tabs.* " _
		&	"FROM tabs " _
		&		"LEFT JOIN tabcategorymap on tabcategorymap.TabID = Tabs.TabID " _
		&	"WHERE " _
		&		sWhere _
		&	";"

	DIM	g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT g_RS.EOF THEN
	
		DIM sFile
		DIM	oLogo
		SET oLogo = g_RS.Fields("Icon")
		
		DIM	oPicture
		SET oPicture = g_RS.Fields("Picture")
		
		DIM	oBackground
		SET oBackground = g_RS.Fields("Background")
		
		DIM	oID
		SET oID = g_RS.Fields("TabID")

		picture_findPath "Categories"
		
		g_RS.MoveFirst		
		DO UNTIL g_RS.EOF
			sFile = recString(oLogo)
			IF "" <> sFile THEN picture_delete sFile
			sFile = recString(oPicture)
			IF "" <> sFile THEN picture_delete sFile
			sFile = recString(oBackground)
			IF "" <> sFile THEN picture_delete sFile

			relations_deleteCategories oID.Value
			
			g_RS.Delete adAffectCurrent
			g_RS.MoveNext
		LOOP

	END IF
	
	Set g_RS.ActiveConnection = Nothing
	g_RS.Close
	SET g_RS = Nothing

END IF

SET g_RS = Nothing

cache_clearFolder "site"

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "tabs.asp"

%>