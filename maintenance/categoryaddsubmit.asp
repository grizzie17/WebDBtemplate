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
<%





DIM sCategoryID
DIM	sName
DIM	sSortName
DIM	sShortName
DIM	sShortDesc
DIM	sTabs

sName = Request("Name")
sSortName = Request("SortName")
sShortName = Request("ShortName")
sShortDesc = Request("ShortDesc")
sTabs = Request("Tabs")



		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM categories " _
			&	"WHERE " _
			&	" CategoryID=0" _
			&	";"
	
DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )

		g_RS.AddNew
		
		g_RS.Fields("Name").Value = sName
		g_RS.Fields("SortName").Value = fieldString(sSortName)
		g_RS.Fields("ShortName").Value = fieldString(sShortName)
		g_RS.Fields("ShortDesc").Value = sShortDesc
		g_RS.Fields("Disabled").Value = 0
		g_RS.Update
		sCategoryID = g_RS.Fields("CategoryID")

		SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

commitTabs CLNG(sCategoryID), sTabs

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "categories.asp"

%>