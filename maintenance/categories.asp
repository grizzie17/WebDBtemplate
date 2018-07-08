<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True

%>
<!--#include virtual="/scripts/config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include virtual="/scripts/include_db_begin.asp"-->
<!--#include virtual="/scripts/htmlformat.asp"-->
<%





DIM	bEditNav
bEditNav = authHasAccess("NAVm")





%><html>

<head>
<meta name="navigate" content="tab">
<title>Categories</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="theme/style.css" type="text/css">


</head>

<body bgcolor="#FFFFFF">
<!--#include file="scripts\index_include.asp"-->

<div style="text-align: right">Categories are used to group pages/articles.&nbsp; 
	Each page must be assigned to a single category</div>

<%
IF bEditNav THEN
%>

<p><a href="categoryadd.asp">Add New Category</a></p>
<%
END IF


	DIM	g_RS
	DIM	oTabRS
	DIM	oDC
	

	DIM	sSelect
	DIM	sTabSelect
	sSelect = "SELECT " _
			&		"c.CategoryID, " _
			&		"c.Name, " _
			&		"c.SortName, " _
			&		"c.ShortName, " _
			&		"c.ShortDesc, " _
			&		"c.Disabled, " _
			&		"COUNT(p.Category) AS PageCount " _
			&	"FROM " _
			&		"categories AS c " _
			&		"LEFT OUTER JOIN " _
			&			"pages as p " _
			&			"ON p.Category = c.CategoryID " _
			&	"WHERE " _
			&		"SUBSTRING( c.Name, 1, 1 ) != '@' " _
			&	"GROUP BY " _
			&		"p.Category, c.Name, c.CategoryID, c.SortName, c.ShortName, c.ShortDesc, c.Disabled " _
			&	"ORDER BY " _
			&		"c.Name " _
			&	";"
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN
	

		DIM	oCategoryID
		DIM	oName
		DIM	oSortName
		DIM	oShortName
		DIM	oShortDesc
		DIM	oPageCount
		DIM	oDisabled

		DIM	nCategoryID
		DIM	sName
		DIM	sSortName
		DIM	sShortName
		DIM	sShortDesc
		DIM	nPageCount
		DIM	bDisabled

		DIM	bDeleteButton
		bDeleteButton = FALSE

		DIM	oTabName
		DIM	oTabDisabled
		DIM	sTabs
		
		DIM	sTemp
		DIM	sStyle
		
		SET oCategoryID = g_RS.Fields("CategoryID")
		SET	oName = g_RS.Fields("Name")
		SET oSortName = g_RS.Fields("SortName")
		SET oShortName = g_RS.Fields("ShortName")
		SET	oShortDesc = g_RS.Fields("ShortDesc")
		SET oPageCount = g_RS.Fields("PageCount")
		SET	oDisabled = g_RS.Fields("Disabled")

%>
<form method="get" action="categorydelete.asp">
<table cellspacing="0" cellpadding="2" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
<tr>
<th><%IF bEditNav THEN%>Edit<%END IF%></th>
<th>Category</th>
<th>Sort</th>
<th>Short</th>
<th>Description</th>
<th>Tabs</th>
<th>Disabled</th>
<th>Pages</th>
<th><%IF bEditNav THEN%>Del<%END IF%></th>
<th>#</th>
</tr>
<%
		g_RS.MoveFirst
		DO UNTIL g_RS.EOF
			nCategoryID = recNumber(oCategoryID)
			sName = TRIM(recString(oName))
			sSortName = TRIM(recString(oSortName))
			sShortName = TRIM(recString(oShortName))
			sShortDesc = TRIM(recString(oShortDesc))
			nPageCount = recNumber(oPageCount)
			bDisabled = recBool(oDisabled)
			IF 0 <> nCategoryID THEN
				sStyle = ""
				IF bDisabled THEN
					sStyle = "style=""color:#CCCCCC"""
				END IF
		%>
		<tr>
		<td valign="top" align="center">
		<%
		IF bEditNav THEN
		%>
		<a href="categoryedit.asp?Category=<%=nCategoryID%>">
        <img src="images/edit.gif" border="0"></a>
		<%
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF "" <> sName THEN
			Response.Write Server.HTMLEncode(sName)
		ELSE
			Response.Write "-none-"
		END IF
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF "" <> oSortName THEN
			Response.Write Server.HTMLEncode(sSortName)
		END IF
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF "" <> sShortName THEN
			Response.Write Server.HTMLEncode(sShortName)
		END IF
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
			IF "" <> sShortDesc THEN
				Response.Write htmlFormatCRLF(sShortDesc)
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
		sTabSelect = "" _
				&	"SELECT " _
				&		"tabs.Name AS Name, " _
				&		"tabs.Disabled AS Disabled " _
				&	"FROM " _
				&		"tabcategorymap " _
				&		"INNER JOIN tabs ON (tabcategorymap.TabID = tabs.TabID) " _
				&	"WHERE " _
				&		"(tabcategorymap.CategoryID = " & nCategoryID & ") " _
				&	"ORDER BY " _
				&		"tabs.Name" _
				&	";"
		
		SET oTabRS = dbQueryRead( g_DC, sTabSelect )
		IF 0 < oTabRS.RecordCount THEN
			SET	oTabName = oTabRS.Fields("Name")
			SET	oTabDisabled = oTabRS.Fields("Disabled")
			sTabs = ""
			oTabRS.MoveFirst
			DO UNTIL oTabRS.EOF
				sTabs = sTabs & ", "
				IF isTrue(oTabDisabled) THEN
					sTabs = sTabs & "<font color=""#CCCCCC"">" & Server.HTMLEncode(recString(oTabName)) & "</font>"
				ELSE
					sTabs = sTabs & Server.HTMLEncode(recString(oTabName))
				END IF
				oTabRS.MoveNext
			LOOP
			Response.Write MID(sTabs,2)
		END IF
		oTabRS.Close
		SET oTabRS = Nothing

		%> &nbsp;
		</td>
		<td valign="top" align="center">
		<%
				IF bDisabled THEN
					%><span style="font-size: smaller">Disabled</span><%
				END IF
		%></td>
		<td valign="top" align="right">
		<a href="pages.asp?Category=<%=nCategoryID%>"><%=nPageCount%> <img src="images/edit.gif" border="0"></a>
		</td>
		<td align="center">
		<%
				IF 0 < nPageCount THEN
		%>
					&nbsp;-&nbsp;
		<%
				ELSEIF bEditNav THEN
					bDeleteButton = TRUE
		%>
		&nbsp;<input type="checkbox" name="category" value="<%=nCategoryID%>">&nbsp;
		<%
				END IF
		%>
		</td>
		<td valign="top" align="right" style="color:#CCCCCC"><%=nCategoryID%>&nbsp;</td>
			
		</tr>
		<%
			END IF
			g_RS.MoveNext
		LOOP
		
		g_RS.Close
		
		IF bDeleteButton THEN
%>
	<tr>
		<td colspan="7">&nbsp;</td>
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
<p>No Categories Exist</p>
<%
	END IF
	


%>
<p><a href="default.asp">Back to Maintenance Page</a></p>
<!--#include virtual="/scripts/byline.asp"-->

</body>

</html>
<!--#include virtual="/scripts/include_db_end.asp"-->