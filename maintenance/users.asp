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
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<title>Users Maintenance</title>
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<style type="text/css">

.DateModified, .DateCreated
{
	font-size: xx-small;
	font-family: sans-serif;
}




</style>

</head>

<body>
<!--#include file="scripts/index_include.asp"-->

<h1>Users</h1>
<p><a href="usersadd.asp">Add New User</a></p>
<%

	DIM	g_RS
	DIM	oCatRS
	
	DIM	sSelect
	DIM	sSelectCats

	sSelect = "SELECT " _
			&		"* " _
			&	"FROM " _
			&		"authorizeusers " _
			&	"ORDER BY " _
			&		"Username " _
			&	";"
	
	SET g_RS = dbQueryRead( g_DC, sSelect )
	
	
	IF NOT g_RS.EOF THEN
	
		DIM	oRID
		DIM	oUsername
		DIM	oFirstname
		DIM	oLastname
		DIM	oEmail
		DIM	oDescription
		DIM	oDateCreated
		DIM oDateModified
		DIM	oDisabled

		DIM	bDeleteButton
		bDeleteButton = FALSE

		DIM	sCategories
		DIM	oCatName
		DIM	oCatDisabled
		
		DIM	sTemp
		DIM	sStyle
		
		SET oRID = g_RS.Fields("RID")
		SET	oUsername = g_RS.Fields("Username")
		SET oFirstname = g_RS.Fields("Firstname")
		SET oLastname = g_RS.Fields("Lastname")
		SET oEmail = g_RS.Fields("Email")
		SET	oDescription = g_RS.Fields("Description")
		SET oDateCreated = g_RS.Fields("DateCreated")
		SET oDateModified = g_RS.Fields("DateModified")
		SET	oDisabled = g_RS.Fields("Disabled")
%>
<form method="get" action="usersdelete.asp">
<table cellspacing="0" cellpadding="2" border="1" style="border-collapse: collapse; border-color:#CCCCCC;">
<tr>
<th>Edit</th>
<th>User</th>
<th>Name</th>
<th>Email</th>
<th>Groups</th>
<th>Dates</th>
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
		%>
		<tr>
		<td valign="top" align="center">
		<a href="usersedit.asp?user=<%=oRID%>">
        <img src="images/edit.gif" border="0" alt="Edit"></a>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		sTemp = recString(oUsername)
		IF "" <> sTemp THEN
			Response.Write Server.HTMLEncode(sTemp)
		ELSE
			Response.Write "-none-"
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		sTemp = TRIM(recString(oFirstname) & " " & recString(oLastname))
		IF "" <> sTemp THEN
			Response.Write Server.HTMLEncode(sTemp)
		ELSE
			Response.Write "-none-"
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
		sTemp = recString(oEmail)
		IF "" <> sTemp THEN
			Response.Write Server.HTMLEncode(sTemp)
		END IF
		%>
		&nbsp;</td>
		<td valign="top" <%=sStyle%>>
		<%
		sSelectCats = "" _
			&	"SELECT " _
			&		"* " _
			&	"FROM authorizegroup " _
			&		"INNER JOIN authorizegroupmap " _
			&			"ON authorizegroupmap.GroupID = authorizegroup.RID " _
			&	"WHERE " _
			&		"authorizegroupmap.UserID = " & oRID & " " _
			&	"ORDER BY " _
			&		dbQ("Name") & " " _
			&	";"
		SET oCatRS = dbQueryRead( g_DC, sSelectCats )
		IF NOT oCatRS IS Nothing THEN
			sTemp = ""
			DO UNTIL oCatRS.EOF
				sTemp = sTemp & ", " & recString(oCatRS.Fields("Name"))
				oCatRS.MoveNext
			LOOP
			sTemp = MID(sTemp, 2)
			Response.Write sTemp
			oCatRS.Close
			SET oCatRS = Nothing
		END IF
		%>
		</td>
		<td valign="top" <%=sStyle%>>
		<%
				IF NOT ISNULL(oDateCreated) THEN
					Response.Write "<div class=""DateCreated"">Cre: " & recDate(oDateCreated) & "</div>"
				END IF
				IF NOT ISNULL(oDateModified) THEN
					Response.Write "<div class=""DateModified"">Mod: " & recDate(oDateModified) & "</div>"
				END IF

		%>
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
				bDeleteButton = TRUE
		%>
		&nbsp;<input type="checkbox" name="tab" value="<%=oRID%>">&nbsp;
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
		<td colspan="2"><input type="submit" value="Delete" name="DeleteUsers"></td>
	</tr>
<%
		END IF
%>
</table>
</form>
<%
	ELSE
%>
<p>No Users Exist</p>
<%
	END IF
	



%>
<p><a href="./">Back to Maintenance Page</a></p>


<!--#include file="scripts\byline.asp"-->

</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
