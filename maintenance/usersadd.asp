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






%><html>

<head>
<title>Add User</title>
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

<body bgcolor="#FFFFFF" id="pageAdd">

<h1>Add User</h1>
<%

DIM	sError
sError = Request("error")
IF "" <> sError THEN
	IF ISNUMERIC(sError) THEN
		sError = authErrorString( CLNG(sError) )
	END IF
	Response.Write "<p style=""color:red"">" & Server.HTMLEncode(sError) & "</p>" & vbCRLF
END IF


DIM	sSelect

DIM	oRS

DIM	nMxsFirstname
DIM	nMxsLastname
DIM	nMxsUsername
DIM	nMxsPassword
DIM	nMxsEmail
DIM	nMxsDescription

nMxsFirstname = dbGetFieldSize(g_DC, "authorizeusers", "Firstname")
nMxsLastname = dbGetFieldSize(g_DC, "authorizeusers", "Lastname")
nMxsUsername = dbGetFieldSize(g_DC, "authorizeusers", "Username")
nMxsPassword = dbGetFieldSize(g_DC, "authorizeusers", "Password")
nMxsEmail = dbGetFieldSize(g_DC, "authorizeusers", "Email")
nMxsDescription = dbGetFieldSize(g_DC, "authorizeusers", "Description")

SET oRS = Nothing

DIM	f
DIM	l
DIM	u
DIM	m
DIM	d
f = Request("f")
l = Request("l")
u = Request("u")
m = Request("m")
d = Request("d")



%>
<form method="POST" action="usersaddsubmit.asp" name="EditForm" onsubmit="return validateForm()">
	<div align="center">
				<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
					<tr>
						<td align="right">Username</td>
						<td>
						<input type="text" name="Username" size="<%=min(20,nMxsUsername)%>" maxlength="<%=nMxsUsername%>" value="<%=u%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Firstname</td>
						<td>
						<input type="text" name="Firstname" size="<%=min(20,nMxsFirstname)%>" maxlength="<%=nMxsFirstname%>" value="<%=f%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Lastname</td>
						<td>
						<input type="text" name="Lastname" size="<%=min(20,nMxsLastname)%>" maxlength="<%=nMxsLastname%>" value="<%=l%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Password</td>
						<td>
						<input name="Password" size="<%=min(20,nMxsPassword)%>" class="required"></td>
					</tr>
					<tr>
						<td align="right">Email</td>
						<td>
						<input type="text" name="Email" size="<%=min(30,nMxsEmail)%>" maxlength="<%=nMxsEmail%>" value="<%=m%>" class="required">
						</td>
					</tr>
					<tr>
						<td align="right">Description</td>
						<td>
						<input type="text" name="Description" size="<%=min(40,nMxsDescription)%>" maxlength="<%=nMxsDescription%>" value="<%=Server.HTMLEncode(d)%>">
						</td>
					</tr>
				</table>
		<p><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>

<!--#include file="scripts\byline.asp"-->


</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->