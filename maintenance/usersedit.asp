<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\vbscriptex.asp"-->
<%


FUNCTION checkTrue( o )
	checkTrue = ""
	IF recBool( o ) THEN
		checkTrue = "checked"
	END IF
END FUNCTION






DIM	sUserID
sUserID = Request("user")



%><html>

<head>
<title>Edit User</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script type="text/javascript" language="JavaScript" src="scripts/formvalidate.js"></script>
<script type="text/javascript" language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "users.asp";
}

function validateForm()
{
	return validateRequired( document.EditForm );
}


//-->
</script>
</head>

<body bgcolor="#FFFFFF" id="pageEdit">

<h1>Edit User</h1>
<%

DIM	sError
sError = Request("error")
IF "" <> sError THEN
	IF ISNUMERIC(sError) THEN
		sError = authErrorString( CLNG(sError) )
	END IF
	Response.Write "<p style=""text-align:center;color:red"">" & Server.HTMLEncode(sError) & "</p>" & vbCRLF
END IF


DIM	sSelect
sSelect = "" _
	&	"SELECT " _
	&		"* " _
	&	"FROM " _
	&		"authorizeusers " _
	&	"WHERE " _
	&		"RID=" & sUserID & " " _
	&	";"

DIM	oRS
SET oRS = dbQueryRead( g_DC, sSelect )

DIM	nMxsFirstname
DIM	nMxsLastname
DIM	nMxsUsername
DIM	nMxsPassword
DIM	nMxsEmail
DIM	nMxsDescription

nMxsFirstname = oRS.Fields("Firstname").DefinedSize
nMxsLastname = oRS.Fields("Lastname").DefinedSize
nMxsUsername = oRS.Fields("Username").DefinedSize
nMxsPassword = oRS.Fields("Password").DefinedSize
nMxsEmail = oRS.Fields("Email").DefinedSize
nMxsDescription = oRS.Fields("Description").DefinedSize

DIM	x
DIM	f
DIM	l
DIM	u
DIM	p
DIM	m
DIM	d
x = checkTrue(oRS.Fields("Disabled"))
f = recString(oRS.Fields("Firstname"))
l = recString(oRS.Fields("Lastname"))
u = recString(oRS.Fields("Username"))
p = recString(oRS.Fields("Password"))
m = recString(oRS.Fields("Email"))
d = recString(oRS.Fields("Description"))

oRS.Close
SET oRS = Nothing


%>
<form method="POST" action="userseditsubmit.asp" name="EditForm" onsubmit="return validateForm()">
	<input type="hidden" name="UserID" value="<%=sUserID%>">
	<div align="center">
	<table>
	<tr>
	<td>
				<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
					<tr>
						<td align="right">&nbsp;</td>
						<td>
						<label id="Disabled">
						<input id="Disabled" name="Disabled" type="checkbox" value="Yes" <%=x%>> 
						Disabled</label></td>
					</tr>
					<tr>
						<td align="right">Username</td>
						<td>
						<input type="text" name="Username" size="<%=min(20,nMxsUsername)%>" maxlength="<%=nMxsUsername%>" value="<%=Server.HTMLEncode(u)%>" disabled="disabled"></td>
					</tr>
					<tr>
						<td align="right">Firstname</td>
						<td>
						<input type="text" name="Firstname" size="<%=min(20,nMxsFirstname)%>" maxlength="<%=nMxsFirstname%>" value="<%=Server.HTMLEncode(f)%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Lastname</td>
						<td>
						<input type="text" name="Lastname" size="<%=min(20,nMxsLastname)%>" maxlength="<%=nMxsLastname%>" value="<%=Server.HTMLEncode(l)%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Password</td>
						<td>
						<input type="text" name="Password" size="<%=min(20,nMxsPassword)%>" maxlength="<%=nMxsPassword%>" value="<%=Server.HTMLEncode(p)%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Email</td>
						<td>
						<input type="text" name="Email" size="<%=min(30,nMxsEmail)%>" maxlength="<%=nMxsEmail%>" value="<%=Server.HTMLEncode(m)%>" class="required">
						</td>
					</tr>
					<tr>
						<td align="right">Description</td>
						<td>
						<input type="text" name="Description" size="<%=min(40,nMxsDescription)%>" maxlength="<%=nMxsDescription%>" value="<%=Server.HTMLEncode(d)%>">
						</td>
					</tr>
				</table>
	</td>
	<td valign="top">
<%

sSelect = "" _
	&	"SELECT " _
	&		"authorizegroup.RID AS xGroupID, " _
	&		"[Name] AS xGroupName, " _
	&		"authorizegroupmap.UserID AS xUserID " _
	&	"FROM authorizegroup " _
	&		"LEFT JOIN authorizegroupmap " _
	&			"ON authorizegroupmap.GroupID = authorizegroup.RID " _
	&	"WHERE " _
	&		"authorizegroupmap.UserID = " & sUserID & " " _
	&	"UNION " _
	&	"SELECT " _
	&		"RID AS xGroupID, " _
	&		"[Name] AS xGroupName, " _
	&		"0 AS xUserID " _
	&	"FROM " _
	&		"authorizegroup " _
	&	"ORDER BY " _
	&		"xGroupName, xUserID Desc " _
	&	";"

SET oRS = dbQueryRead( g_DC, sSelect )
IF NOT oRS IS Nothing THEN
	DIM	oRID
	DIM	oName
	DIM	oMapUser
	SET oRID = oRS.Fields("xGroupID")
	SET oName = oRS.Fields("xGroupName")
	SET oMapUser = oRS.Fields("xUserID")
	DIM	sLabel
	DIM	sValue
	DIM	sName
	DIM	sPrevName
	
	sPrevName = "******"

	DIM	i
	i = 0
	oRS.MoveFirst
	DO UNTIL oRS.EOF
		sName = Server.HTMLEncode(recString(oName))
		IF sPrevName <> sName THEN
			i = i + 1
			sValue = recNumber(oRID)
			sLabel = "Group-" & i
			Response.Write "<div class=""checkgroups"">"
			Response.Write "<input type=""checkbox"" "
			Response.Write "name=""" & sLabel & """ id=""" & sLabel & """ "
			IF 0 < recNumber(oMapUser) THEN
				Response.Write "checked=""checked"" "
			END IF
			Response.Write "value=""" & sValue & """"
			Response.Write ">"
			Response.Write "<label for=""" & sLabel & """> " & sName & "</label>"
			'Response.Write " " & recNumber(oMapUser)
			Response.Write "</div>"
			sPrevName = sName
		END IF
		oRS.MoveNext
	LOOP
	oRS.Close
	SET oRS = Nothing
END IF

%>
	</td>
	</tr>
	<tr><td><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></td><td></td></tr>
	</table>
	</div>
</form>

<!--#include file="scripts\byline.asp"-->
</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->