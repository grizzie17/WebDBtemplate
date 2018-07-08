<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%

FUNCTION isTrue( o )
	isTrue = 0
	IF "ON" = UCASE(o) THEN
		isTrue = -1
	END IF
END FUNCTION




picture_createObject

DIM	RID
DIM	sCategory
DIM	sProdID
DIM	sName
DIM	sSortName
DIM	sFormat
DIM	sTitle
DIM	sDesc
DIM	sShortDesc
DIM	sBody

DIM	sQualCategory
sQualCategory = Requestor("QualCategory")


sCategory = Requestor("Category")
sName = Requestor("PageName")
sSortName = Requestor("SortName")
'sFormat = Requestor("Format")
sTitle = Requestor("Title")
sDesc = Requestor("Description")
'sBody = Requestor("BodyText")





		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM pages " _
			&	"WHERE " _
			&	" RID=0" _
			&	";"
	
DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )





		g_RS.AddNew
		
		g_RS.Fields("PageName").Value = sName
		g_RS.Fields("SortName").Value = sSortName
		'g_RS.Fields("Format").Value = sFormat
		g_RS.Fields("Title").Value = sTitle
		g_RS.Fields("Description").Value = sDesc
		'g_RS.Fields("Body").Value = sBody
		g_RS.Fields("DateCreated").Value = Now
		g_RS.Fields("DateModified").Value = Now
		g_RS.Fields("Category").Value = sCategory
		g_RS.Fields("Disabled").Value = 0
		
		g_RS.Update
		RID = g_RS.Fields("RID").Value
		
		SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "pageedit.asp?page=" & RID & "&category=" & sQualCategory

%>