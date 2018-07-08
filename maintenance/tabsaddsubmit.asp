<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_relations.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%

FUNCTION isTrue( o )
	isTrue = 0
	IF "ON" = o THEN
		isTrue = -1
	END IF
END FUNCTION







DIM	sTabID
DIM	sName
DIM	sSortName
DIM	sDescription
DIM	sCategories

sName = Request("Name")
sSortName = Request("SortName")
sDescription = Request("Description")
sCategories = Request("Categories")




		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM tabs " _
			&	"WHERE " _
			&	" TabID=0" _
			&	";"
	
DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )

		g_RS.AddNew
		
		g_RS.Fields("Name").Value = sName
		g_RS.Fields("SortName").Value = sSortName
		g_RS.Fields("Description").Value = sDescription
		g_RS.Fields("Disabled").Value = 0
		g_RS.Update
		sTabID = g_RS.Fields("TabID")

		'SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

commitCategories CLNG(sTabID), sCategories

cache_clearFolder "site"


%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "tabs.asp"

%>