<%@ LANGUAGE="VBSCRIPT" %>
<%
OPTION EXPLICIT

Response.Expires=0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
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
		&		"Name " _
		&	"FROM " _
		&		"Categories " _
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
<meta name="navigate" content="!tab">
<title>Products</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../theme/style.css" type="text/css">

<script language="JavaScript">
<!--


//-->
</script>

</head>

<body bgcolor="#FFFFFF">
<!--#include file="scripts/index_include.asp"-->





<p><a href="productadd.asp?category=<%=sCategory%>&brand=<%=sBrand%>">Add New 
<%
IF "" <> sCategory THEN
	Response.Write sCategoryName
END IF
%>
Page</a></p>
<%

DIM	sWhere
IF "" <> sCategory THEN
	sWhere = "WHERE Category = " & sCategory & " "
ELSEIF "" <> sBrand THEN
	sWhere = "WHERE Brand = " & sBrand & " "
ELSE
	sWhere = " "
END IF
	

	sSelect = "SELECT " _
			&		"Products.RID, " _
			&		"Products.ProdID, " _
			&		"Products.Name, " _
			&		"Products.ShortDesc, " _
			&		"Products.UnitPrice, " _
			&		"Products.Disabled, " _
			&		"Brands.Name AS BrandName, " _
			&		"Categories.Name AS CatName " _
			&	"FROM " _
			&		"(Products " _
			&		"LEFT JOIN Categories ON Categories.CategoryID = Products.Category) " _
			&		"LEFT JOIN Brands ON Brands.BrandID = Products.Brand " _
			&	sWhere _
			&	"ORDER BY " _
			&		"Categories.Name, IIF( ISNULL(Products.SortName),'',Products.SortName)+Products.Name" _
			&	";"
	
	
	DIM	g_RS
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN

		DIM	sCatName
		sCatName = ""

		DIM	oID
		DIM	oRID
		DIM	oName
		DIM	oDesc
		DIM	oBrand
		DIM	oPrice
		DIM	oProdCount
		DIM	oDisabled
		DIM	oCatName

		DIM	bDeleteButton
		bDeleteButton = FALSE

		DIM	sTemp
		DIM	sStyle
		
		SET oRID = g_RS.Fields("RID")
		SET oID = g_RS.Fields("ProdID")
		SET	oName = g_RS.Fields("Name")
		SET	oDesc = g_RS.Fields("ShortDesc")
		SET oBrand = g_RS.Fields("BrandName")
		SET oPrice = g_RS.Fields("UnitPrice")
		SET	oDisabled = g_RS.Fields("Disabled")
		SET oCatName = g_RS.Fields("CatName")
%>
<form method="get" action="productdelete.asp">
<input type="hidden" name="QualCategory" value="<%=sCategory%>">
<input type="hidden" name="QualBrand" value="<%=sBrand%>">
<table cellspacing="0" cellpadding="2" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
<tr>
<th>Edit</th>
<th>Name</th>
<th>Description</th>
<th>Brand</th>
<th>P-ID</th>
<th>Disabled</th>
<th>Del</th>
<th>#</th>
</tr>
<%
		DO UNTIL g_RS.EOF
			IF 0 <> oRID THEN
				sStyle = ""
				IF isTrue(oDisabled) THEN
					sStyle = "style=""color:#CCCCCC"""
				END IF
			IF sCatName <> recString(oCatName) THEN
				sCatName = recString(oCatName)
				IF "" <> sCatName THEN
%>
				<tr>
				<th align="left" colspan="7" bgcolor="#DDDDDD">
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
		<a href="productedit.asp?category=<%=sCategory%>&brand=<%=sBrand%>&product=<%=oRID%>">
        <img src="../images/edit.gif" border="0"></a>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF NOT ISNULL(oName) THEN
			Response.Write Server.HTMLEncode(recString(oName))
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
			Response.Write Server.HTMLEncode(recString(oDesc))
			IF FALSE THEN
		%>
		&nbsp;
		<%
			END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
			Response.Write Server.HTMLEncode(recString(oBrand))
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF NOT ISNULL(oName) THEN
			Response.Write Server.HTMLEncode(recString(oID))
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
		<td valign="top" align="center">
		<%
				IF isTrue(oDisabled) THEN
					%><span style="font-size: smaller">Disabled</span><%
				END IF
		%></td>
		<td align="center" valign="top">
		<%
			bDeleteButton = TRUE
		%>
		&nbsp;<input type="checkbox" name="product" value="<%=oRID%>">&nbsp;
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
<p>No Pages Exist</p>
<%
	END IF
	



IF "" <> sCategory THEN
%>
<p><a href="../categories.asp">Back to Categories Page</a></p>
<%
ELSEIF "" <> sBrand THEN
%>
<p><a href="brands.asp">Back to Brands Page</a></p>
<%
ELSE
%>
<p><a href="../">Back to Content Maintenance Page</a></p>
<%
END IF
%>

<!--webbot bot="Include" U-Include="../../_private/byline.htm" TAG="BODY" startspan -->

<script language="JavaScript">
<!--

function makeByLine()
{
	document.write( '<' + 'script language="JavaScript" src="http://' );
	if ( "localhost" == location.hostname )
	{
		document.writeln( 'localhost/BearConsultingGroup' );
	}
	else
	{
		document.writeln( 'BearConsultingGroup.com' );
	}
	document.writeln( '/designby_small.js"></' + 'script>' );
}
makeByLine()

//-->
</script>
<!--script language="JavaScript" src="http://BearConsultingGroup.com/designbyadvert.js"></script-->

<!--webbot bot="Include" i-checksum="20551" endspan -->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->