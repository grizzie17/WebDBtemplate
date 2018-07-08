<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<%

FUNCTION isTrue( o )
	isTrue = 0
	IF "ON" = o THEN
		isTrue = -1
	END IF
END FUNCTION











DIM	sName
DIM	sDesc
DIM	sWebsite
DIM	sRating

sName = Request("Name")
sDesc = Request("Desc")
sWebsite = Request("Website")
sRating = Request("Rating")




		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM Brands " _
			&	"WHERE " _
			&	" BrandID=0" _
			&	";"
	
DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )

		g_RS.AddNew
		
		g_RS.Fields("Name").Value = sName
		g_RS.Fields("Description").Value = sDesc
		g_RS.Fields("Website").Value = sWebsite
		g_RS.Fields("Rating").Value = fieldNumber( sRating )
		g_RS.Fields("Disabled").Value = 0
		g_RS.Update

		SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "brands.asp"

%>