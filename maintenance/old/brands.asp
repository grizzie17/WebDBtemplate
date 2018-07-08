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












%><html>

<head>
<meta name="navigate" content="!tab">
<title>Brands</title>
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




<p><a href="brandadd.asp">Add New Brand</a></p>
<%

	

	DIM	sSelect
	sSelect = "SELECT " _
			&		"BrandID, " _
			&		"Name, " _
			&		"Description, " _
			&		"Website, " _
			&		"Rating, " _
			&		"Logo, " _
			&		"Disabled, " _
			&		"(SELECT Count(*) FROM Products WHERE BrandID=Products.Brand) AS BrandCount " _
			&	"FROM " _
			&		"Brands " _
			&	"ORDER BY " _
			&		"Name" _
			&	";"
	
	
	DIM	g_RS
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN
	
		DIM	oID
		DIM	oName
		DIM	oDesc
		DIM	oWebsite
		DIM	oRating
		DIM	oLogo
		DIM	oDisabled
		DIM	oBrandCount
		
		DIM	sPictureName
		
		DIM	bDeleteButton
		bDeleteButton = FALSE
		
		DIM	sTemp
		DIM	sStyle
		
		SET oID = g_RS.Fields("BrandID")
		SET	oName = g_RS.Fields("Name")
		SET	oDesc = g_RS.Fields("Description")
		SET	oWebsite = g_RS.Fields("Website")
		SET oRating = g_RS.Fields("Rating")
		SET oLogo = g_RS.Fields("Logo")
		SET	oDisabled = g_RS.Fields("Disabled")
		SET oBrandCount = g_RS.Fields("BrandCount")
%>
<form method="get" action="branddelete.asp">
<table cellspacing="0" cellpadding="2" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
<tr>
<th>Edit</th>
<th>Brand</th>
<th>Description</th>
<th>Website</th>
<th>Rating</th>
<th>Disabled</th>
<th>Prods</th>
<th>Del</th>
<th>#</th>
</tr>
<%
		DO UNTIL g_RS.EOF
			IF 0 <> oID THEN
				sStyle = ""
				IF isTrue(oDisabled) THEN
					sStyle = "style=""color:#CCCCCC"""
				END IF
		%>
		<tr>
		<td valign="top" align="center">
		<a href="brandedit.asp?brand=<%=oID%>">
        <img src="../images/edit.gif" border="0"></a>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		IF NOT ISNULL(oName) THEN
			Response.Write Server.HTMLEncode(recString(oName))
		ELSE
			Response.Write "-none-"
		END IF
		sPictureName = recString(oLogo)
		IF "" <> sPictureName THEN
			Response.Write "<br>"
			Response.Write "<img src=""picture.asp?table=brands&name=" & sPictureName & """>" & vbCRLF
		END IF
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
			Response.Write htmlFormatCRLF(recString(oDesc))
			IF FALSE THEN
		%>
		&nbsp;
		<%
			END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
			Response.Write Server.HTMLEncode(recString(oWebsite))
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
			Response.Write Server.HTMLEncode(recNumber(oRating))
		%>
		&nbsp;</td>
		<td valign="top" align="center">
		<%
				IF isTrue(oDisabled) THEN
					%><span style="font-size: smaller">Disabled</span><%
				END IF
		%></td>
		<td valign="top" align="right"><a href="products.asp?brand=<%=oID%>"><%=oBrandCount%><img src="../images/edit.gif" border="0"></a>&nbsp;</td>
		<td align="center">
		<%
				IF 0 < oBrandCount THEN
		%>
					&nbsp;-&nbsp;
		<%
				ELSE
					bDeleteButton = TRUE
		%>
		&nbsp;<input type="checkbox" name="brand" value="<%=oID%>">&nbsp;
		<%
				END IF
		%>
		</td>
		<td valign="top" align="right" style="color:#CCCCCC"><%=oID%>&nbsp;</td>
			
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
		<td colspan="2"><input type="submit" value="Delete" name="DeleteBrands"></td>
	</tr>
<%
		END IF
%>
</table>
</form>
<%
	ELSE
%>
<p>No Brands Exist</p>
<%
	END IF
	




%>
<p><a href="../default.asp">Back to Maintenance Page</a></p>
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