<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\configdb.asp"-->
<%


    DIM sTabName
    sTabName = Request("name")
    IF "" = sTabName THEN
        sTabName = Request("tab")
    END IF

    sTabName = REPLACE(sTabName, "'", "")
    sTabName = REPLACE(sTabName, "-", "")


DIM	nTabID
nTabID = 0

IF "" <> g_sSiteTabNewsletters THEN

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&	"	TabID " _
		&	"FROM " _
		&	"	tabs " _
		&	"WHERE " _
		&	"	tabs.Name LIKE '" & sTabName & "' " _
		&	";"
	
	DIM	oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		IF NOT oRS.EOF THEN
			nTabID = recNumber(oRS.Fields("TabID"))
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF

END IF

%><!--#include file="scripts\include_db_end.asp"-->
<%

IF 0 < nTabID THEN
	Response.Redirect "Page.asp?tab=" & nTabID
ELSE
	Response.Redirect "./"
END IF


%>
