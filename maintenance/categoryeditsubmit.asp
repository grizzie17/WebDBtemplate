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




DIM	CategoryID
DIM	sName
DIM	sSortName
DIM	sShortName
DIM	sShortDesc
DIM	bDisabled
DIM	sTabs

CategoryID = Request("CategoryID")
sName = Request("Name")
sSortName = Request("SortName")
sShortName = Request("ShortName")
sShortDesc = Request("ShortDesc")
bDisabled = Request("Disabled")
sTabs = Request("Tabs")



DIM	sSelect
sSelect = "SELECT " _
		&	" * " _
		&	"FROM categories " _
		&	"WHERE " _
		&	" CategoryID=" & CategoryID & " " _
		&	";"

DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
		
		IF NOT g_RS.EOF THEN
			
			g_RS.Fields("Name").Value = sName
			g_RS.Fields("SortName").Value = fieldString(sSortName)
			g_RS.Fields("ShortName").Value = fieldString(sShortName)
			g_RS.Fields("ShortDesc").Value = sShortDesc
			g_RS.Fields("Revision").Value = recNumber(g_RS.Fields("Revision")) + 1
			g_RS.Fields("Disabled").Value = fieldBool(bDisabled)
			g_RS.Update
		
		END IF

		'SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

commitTabs CLNG(CategoryID), sTabs


%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "categories.asp"

%>