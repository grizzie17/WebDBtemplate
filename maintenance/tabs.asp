<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\htmlformat.asp"-->
<%


FUNCTION isTrue( o )
	isTrue = FALSE
	IF NOT ISNULL( o ) THEN
		IF CBOOL( o ) THEN
			isTrue = TRUE
		ELSE
			isTrue = FALSE
		END IF
	END IF
END FUNCTION





DIM	bEditNav
bEditNav = authHasAccess("NAVm")





%><html>

<head>
<meta name="navigate" content="tab">
<title>Tabs</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="theme/style.css" type="text/css">

</head>

<body bgcolor="#FFFFFF">
<!--#include file="scripts/index_include.asp"-->
<div style="text-align:right">Tabs identify the top navigation for the website. Categories may be associated to many 
	tabs to display their content.</div>


<%
IF bEditNav THEN
%>
<p><a href="tabsadd.asp">Add New Tab</a></p>
<%
END IF

	DIM	g_RS
	DIM	oCatRS
	
	DIM	sSelect
	DIM	sSelectCats

	sSelect = "SELECT " _
			&		"TabID, " _
			&		"Name, " _
			&		"SortName, " _
			&		"Description, " _
			&		"Picture, " _
			&		"Background, " _
			&		"Disabled " _
			&	"FROM " _
			&		"tabs " _
			&	"ORDER BY " _
			&		"Name" _
			&	";"
	
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN
	
		DIM	oTabID
		DIM	oName
		DIM	oSortName
		DIM	oDescription
		DIM	oPicture
		DIM oBackground
		DIM	oDisabled

		DIM	bDeleteButton
		bDeleteButton = FALSE

		DIM	sCategories
		DIM	oCatName
		DIM	oCatDisabled
		
		DIM	sTemp
		DIM	sStyle
		
		SET oTabID = g_RS.Fields("TabID")
		SET	oName = g_RS.Fields("Name")
		SET oSortName = g_RS.Fields("SortName")
		SET	oDescription = g_RS.Fields("Description")
		SET	oDisabled = g_RS.Fields("Disabled")
%>
<form method="get" action="tabsdelete.asp">
<table cellspacing="0" cellpadding="2" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
<tr>
<th><%IF bEditNav THEN%>Edit<%END IF%></th>
<th>Tab</th>
<th>Sort Name</th>
<th>Description</th>
<th>Categories</th>
<th>Disabled</th>
<th><%IF bEditNav THEN%>Del<%END IF%></th>
<th>#</th>
</tr>
<%
		DO UNTIL g_RS.EOF
			IF 0 <> oTabID THEN
				sStyle = ""
				IF isTrue(oDisabled) THEN
					sStyle = "style=""color:#CCCCCC"""
				END IF
		%>
		<tr>
		<td valign="top" align="center">
		<%
		IF bEditNav THEN
		%>
		<a href="tabsedit.asp?tab=<%=oTabID%>">
        <img src="images/edit.gif" border="0"></a>
        <%
        END IF
        %>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		sTemp = recString(oName)
		IF "" <> sTemp THEN
			Response.Write Server.HTMLEncode(sTemp)
		ELSE
			Response.Write "-none-"
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		sTemp = recString(oSortName)
		IF "" <> sTemp THEN
			Response.Write Server.HTMLEncode(sTemp)
		END IF
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
		sTemp = recString(oDescription)
		IF "" <> sTemp THEN
			Response.Write htmlFormatCRLF(recString(oDescription))
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		sSelectCats = "" _
				&	"SELECT " _
				&		"categories.Name AS Name, " _
				&		"categories.Disabled AS Disabled " _
				&	"FROM " _
				&		"tabcategorymap " _
				&		"INNER JOIN categories ON (tabcategorymap.CategoryID = categories.CategoryID) " _
				&	"WHERE " _
				&		"(tabcategorymap.TabID = " & oTabID & ") " _
				&	"ORDER BY " _
				&		"categories.Name" _
				&	";"
		
		SET oCatRS = dbQueryRead( g_DC, sSelectCats )
		IF 0 < oCatRS.RecordCount THEN
			SET	oCatName = oCatRS.Fields("Name")
			SET	oCatDisabled = oCatRS.Fields("Disabled")
			sCategories = ""
			DO UNTIL oCatRS.EOF
				sCategories = sCategories & ", "
				IF isTrue(oCatDisabled) THEN
					sCategories = sCategories & "<font color=""#CCCCCC"">" & Server.HTMLEncode(oCatName) & "</font>"
				ELSE
					sCategories = sCategories & Server.HTMLEncode(oCatName)
				END IF
				oCatRS.MoveNext
			LOOP
			Response.Write MID(sCategories,2)
		END IF
		oCatRS.Close

		%> &nbsp;
		</td>
		<td valign="top" align="center">
		<%
				IF isTrue(oDisabled) THEN
					%><span style="font-size: smaller">Disabled</span><%
				END IF
		%>
		</td>
		<td align="center">
		<%
			IF bEditNav THEN
				bDeleteButton = TRUE
		%>
		&nbsp;<input type="checkbox" name="tab" value="<%=oTabID%>">&nbsp;
		<%
			END IF
		%>
		</td>
		<td valign="top" align="right" style="color:#CCCCCC"><%=oTabID%>&nbsp;</td>
			
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
		<td colspan="2"><input type="submit" value="Delete" name="DeleteTabs"></td>
	</tr>
<%
		END IF
%>
</table>
</form>
<%
	ELSE
%>
<p>No Tabs Exist</p>
<%
	END IF
	



%>
<p><a href="default.asp">Back to Maintenance Page</a></p>


<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->