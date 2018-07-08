<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Server.ScriptTimeout = 60*30
Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\htmlformat.asp"-->
<%




DIM	sBrand
sBrand = Request("brand")


DIM	sCategory
sCategory = Request("category")

DIM	sCategoryName
sCategoryName = ""

DIM	sSelect

IF "" <> sCategory THEN

	sSelect = "" _
		&	"SELECT " _
		&		dbQ("Name") & " " _
		&	"FROM " _
		&		"categories " _
		&	"WHERE " _
		&		"CategoryID = " & sCategory _
		&	";"

	DIM	oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT Nothing IS oRS THEN
		IF 0 < oRS.RecordCount THEN
			sCategoryName = oRS.Fields("Name").Value
			sCategoryName = Server.HTMLEncode(sCategoryName)
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
ELSE
	sCategoryName = "All"
END IF


%><html>

<head>
<meta name="navigate" content="tab">
<title>Pages</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" href="theme/style.css" type="text/css">

<style type="text/css">

.DateModified
{
	font-size: xx-small;
	font-family: sans-serif;
}




</style>


</head>

<body bgcolor="#FFFFFF">
<!--#include file="scripts/index_include.asp"-->





<p><a href="pageadd.asp?category=<%=sCategory%>&brand=<%=sBrand%>">Add New 
<%
IF "" <> sCategory THEN
	Response.Write sCategoryName
END IF
%>
Page</a></p>
<%

DIM	sWhere
IF "" <> sCategory THEN
	sWhere = "AND category = " & sCategory & " "
ELSEIF "" <> sBrand THEN
	sWhere = "AND brand = " & sBrand & " "
ELSE
	sWhere = " "
END IF
	

	sSelect = "SELECT " _
			&		"pages.RID, " _
			&		"pages.PageName, " _
			&		"pages.Format, " _
			&		"pages.Title, " _
			&		"pages.Description, " _
			&		"pages.DateModified, " _
			&		"pages.Disabled, " _
			&		"categories.Name AS CatName, " _
			&		"schedules.DateBegin, " _
			&		"schedules.DateEnd, " _
			&		"schedules.DateEvent, " _
			&		"(CASE WHEN ((ISNULL(schedules.DateBegin,GETDATE()) <= GETDATE())" _
			&		" AND " _
			&		"(GETDATE() <= ISNULL(schedules.DateEnd,GETDATE()))) THEN 0 ELSE 1 END) AS DateDisabled " _
			&	"FROM " _
			&		"(pages " _
			&		"INNER JOIN categories ON categories.CategoryID = pages.Category) " _
			&		"LEFT JOIN schedules ON schedules.PageID = pages.RID " _
			&	"WHERE SUBSTRING( categories.Name, 1, 1 ) != '@' " _
			&		sWhere _
			&	"ORDER BY " _
			&		"categories.Name, ISNULL( pages.SortName,'')+pages.PageName" _
			&	";"
	
	
	DIM	g_RS
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN

		DIM	sCatName
		sCatName = ""

		DIM	oID
		DIM	oRID
		DIM	oPageName
		DIM	oFormat
		DIM	oTitle
		DIM	oDesc
		DIM	oDateModified
		DIM	oDisabled
		DIM	oCatName
		DIM	oDateBegin
		DIM	oDateEnd
		DIM	oDateEvent
		DIM	oDateDisabled

		DIM	bDeleteButton
		bDeleteButton = FALSE

		DIM	sTemp
		DIM	sStyle
		DIM	sFormat
		
		SET oRID = g_RS.Fields("RID")
		SET	oPageName = g_RS.Fields("PageName")
		SET oFormat = g_RS.Fields("Format")
		SET oTitle = g_RS.Fields("Title")
		SET oDesc = g_RS.Fields("Description")
		SET oDateModified = g_RS.Fields("DateModified")
		SET	oDisabled = g_RS.Fields("Disabled")
		SET oCatName = g_RS.Fields("CatName")
		SET oDateBegin = g_RS.Fields("DateBegin")
		SET oDateEnd = g_RS.Fields("DateEnd")
		SET oDateEvent = g_RS.Fields("DateEvent")
		SET oDateDisabled = g_RS.Fields("DateDisabled")
%>
<form method="get" action="pagedelete.asp">
<input type="hidden" name="QualCategory" value="<%=sCategory%>">
<input type="hidden" name="QualBrand" value="<%=sBrand%>">
<table cellspacing="0" cellpadding="2" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
<tr>
<th>Edit</th>
<th>Name</th>
<th>Format</th>
<th>Title</th>
<th>Description</th>
<th>Schedule</th>
<th>Del</th>
<th>#</th>
</tr>
<%
		g_RS.MoveFirst
		DO UNTIL g_RS.EOF
			IF 0 <> oRID THEN
				sStyle = ""
				IF isTrue(oDisabled)  OR  isTrue(oDateDisabled) THEN
					sStyle = "style=""color:#CCCCCC"""
				END IF
				sFormat = recString(oFormat)
				IF "" = sFormat THEN
					sFormat = "STFT"
				END IF
			IF sCatName <> recString(oCatName) THEN
				sCatName = recString(oCatName)
				IF "" <> sCatName THEN
%>
				<tr>
				<th align="left" colspan="6" bgcolor="#DDDDDD">
				<font size="-1"><%=Server.HTMLEncode(sCatName)%> &nbsp;</font></th>
				<th bgcolor="#DDDDDD"><font color="#999999" size="-1">Del</font></th>
				<th bgcolor="#DDDDDD"><font color="#999999" size="-1">#</font></th>
				</tr>
<%
				END IF
			END IF
%>
		<tr>
		<td valign="top" align="center">
		<a href="pageedit.asp?category=<%=sCategory%>&page=<%=oRID%>">
        <img src="images/edit.gif" border="0"></a>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF NOT ISNULL(oPageName) THEN
			Response.Write Server.HTMLEncode(recString(oPageName))
		ELSE
			Response.Write "-none-"
		END IF
		IF FALSE THEN
		%>
		&nbsp;
		<%
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF NOT ISNULL(oFormat) THEN
			Response.Write Server.HTMLEncode(recString(oFormat))
		ELSE
			Response.Write "-none-"
		END IF
		IF FALSE THEN
		%>
		&nbsp;
		<%
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
			SELECT CASE sFormat
			CASE "STFT"
				Response.Write htmlFormatMeta(recString(oTitle))
			CASE "HTML"
				Response.Write recString(oTitle)
			CASE "TEXT"
				Response.Write Server.HTMLEncode(RecString(oTitle))
			CASE ELSE
				Response.Write Server.HTMLEncode(recString(oTitle))
			END SELECT
			IF FALSE THEN
		%>
		&nbsp;
		<%
			END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
			SELECT CASE sFormat
			CASE "STFT"
				Response.Write htmlFormatMeta(recString(oDesc))
			CASE "HTML"
				Response.Write recString(oDesc)
			CASE "TEXT"
				Response.Write Server.HTMLEncode(RecString(oDesc))
			CASE ELSE
				Response.Write Server.HTMLEncode(recString(oDesc))
			END SELECT
			IF FALSE THEN
		%>
		&nbsp;
		<%
			END IF
		%>
		</td>
		<td valign="top" align="center">
		<%
				IF isTrue(oDisabled) THEN
					%><div style="font-size: smaller">Disabled</div><%
				END IF
				IF NOT ISNULL(oDateEvent) THEN
					Response.Write "<div>Event: " & recDate(oDateEvent) & "</div>"
				END IF
				IF NOT ISNULL(oDateBegin) THEN
					Response.Write "<div>Begin: " & recDate(oDateBegin) & "</div>"
				END IF
				IF NOT ISNULL(oDateEnd) THEN
					Response.Write "<div>End: " & recDate(oDateEnd) & "</div>"
				END IF
				IF NOT ISNULL(oDateModified) THEN
					Response.Write "<div class=""DateModified"">Mod: " & recDate(oDateModified) & "</div>"
				END IF
		%></td>
		<td align="center" valign="top">
		<%
			bDeleteButton = TRUE
		%>
		&nbsp;<input type="checkbox" name="page" value="<%=recNumber(oRID)%>">&nbsp;
		</td>
		<td valign="top" align="right" style="color:#CCCCCC"><%=oRID%>&nbsp;</td>
	
		</tr>
		<%
			END IF
			g_RS.MoveNext
		LOOP
		
		g_RS.Close
		IF bDeleteButton THEN
%>
	<tr>
		<td colspan="6">&nbsp;</td>
		<td colspan="2"><input type="submit" value="Delete" name="DeleteCategories"></td>
	</tr>
<%
		END IF
%>
</table>
</form>
<%
	ELSE
%>
<p>No Pages Exist</p>
<%
	END IF
	



IF "" <> sCategory THEN
%>
<p><a href="categories.asp">Back to Categories Page</a></p>
<%
ELSE
%>
<p><a href="./">Back to Content Maintenance Page</a></p>
<%
END IF
%>

<!--#include file="scripts\byline.asp"-->
</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->