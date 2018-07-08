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
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="navigate" content="!tab">
<title>Galleries</title>
<link rel="stylesheet" href="theme/style.css" type="text/css">
</head>

<body>
<!--#include file="scripts/index_include.asp"-->
<p><a href="galleryadd.asp">Add New Gallery</a></p>
<%


	DIM	g_RS
	DIM	oCatRS
	
	DIM	sSelect
	DIM	sSelectCats

	sSelect = "SELECT " _
			&		"RID, " _
			&		"Title, " _
			&		"EventDate, " _
			&		"Description, " _
			&		"DateCreated, " _
			&		"Disabled " _
			&	"FROM " _
			&		"Gallery " _
			&	"ORDER BY " _
			&		"EventDate" _
			&	";"
	
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN
	
		DIM	oRID
		DIM	oTitle
		DIM	oEventDate
		DIM	oDescription
		DIM	oDateCreated
		DIM	oDisabled

		DIM	bDeleteButton
		bDeleteButton = FALSE
		DIM	sStyle

		DIM	nRID
		DIM	sTitle
		DIM	dEventDate
		DIM	sDescription
		DIM	dDateCreated
		DIM	bDisabled
		
		SET oRID = g_RS.Fields("RID")
		SET oTitle = g_RS.Fields("Title")
		SET oEventDate = g_RS.Fields("EventDate")
		SET oDescription = g_RS.Fields("Description")
		SET oDateCreated = g_RS.Fields("DateCreated")
		SET oDisabled = g_RS.Fields("Disabled")

%>
<form method="get" action="gallerydelete.asp">
<table cellspacing="0" cellpadding="2" border="1" bordercolor="#CCCCCC" style="border-collapse: collapse">
<tr>
<th>Edit</th>
<th>Title</th>
<th>Event Date</th>
<th>Description</th>
<th>Categories</th>
<th>Disabled</th>
<th>Del</th>
<th>#</th>
</tr>
<%

		g_RS.MoveFirst
		DO UNTIL g_RS.EOF
			nRID = recNumber(oRID)
			sTitle = recString(oTitle)
			dEventDate = recDate(oEventDate)
			sDescription = recString(oDescription)
			dDateCreated = recDate(oDateCreated)
			bDisabled = recBool(oDisabled)
			sStyle = ""
			IF bDisabled THEN sStyle = " style=""color:#CCCCCC;"""
			%>
			<tr>
			<%
			' Edit column
			%>
			<td valign="top" align="center">
			<a href="galleryedit.asp?id=<%=nRID%>">
	        <img alt="Edit" src="images/edit.gif" border="0"></a>
	        </td>
			<%
			' Title Column
			%>
			<td <%=sStyle%>>
				<%
				IF "" <> sTitle THEN
					Response.Write Server.HTMLEncode(sTitle)
				ELSE
					Response.Write "-Empty-"
				END IF
				%>
			</td>
			<%
			' Event Date Column
			%>
			<td <%=sStyle%>>
				<%
				Response.Write DateValue(dEventDate)
				%>
			</td>
			<%
			' Description Column
			%>
			<td <%=sStyle%>>
				<%
				IF "" <> sDescription THEN
					Response.Write HtmlFormatCRLF( sDescription )
				ELSE
					Response.Write "&nbsp;"
				END IF
				%>
			</td>
			<%
			' Categories Column
			%>
			<td>
			Future
			</td>
			<%
			' Disabled Column
			%>
			<td align="center">
				<%
				IF bDisabled THEN
					%><span style="font-size: smaller">Disabled</span><%
				END IF
				%>
			</td>
			<%
			
			%>
			<td align="center">
			<%
					bDeleteButton = TRUE
			%>
			&nbsp;<input type="checkbox" name="tab" value="<%=nRID%>">&nbsp;
			</td>
			<td valign="top" align="right" style="color:#CCCCCC"><%=nRID%>&nbsp;</td>
			</tr>
			<%
			g_RS.MoveNext
		LOOP
		IF bDeleteButton THEN
%>
	<tr>
		<td colspan="6">&nbsp;</td>
		<td colspan="2"><input type="submit" value="Delete" name="DeleteGalleries"></td>
	</tr>
<%
		END IF
%>
</table>
</form>
<%
		
		g_RS.Close
		SET g_RS = Nothing
	ELSE
%>
<p>No Galleries Exist</p>
<%
	END IF


%>
<p><a href="default.asp">Back to Maintenance Page</a></p>

<!--webbot bot="Include" U-Include="../_private/byline.htm" TAG="BODY" startspan -->

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
