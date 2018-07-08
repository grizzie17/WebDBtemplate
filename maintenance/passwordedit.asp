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
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="noindex">
<script language="JavaScript">
<!--
function doLoad()
{
	document.pwForm.pw1.focus();
}

function checkRequired( theForm )
{
    var bMissing = false;
    
    if ( 0 == theForm.pw1.value.length 
    		||  0 == theForm.pw2.value.length )
    {
        alert( "Both passwords are required.\n"
                + "Please complete them and Submit again.");
        
        // false causes the form submission to be canceled
        return false;
    }
    
    if ( theForm.pw1.value != theForm.pw2.value )
    {
        alert( "Both Password values must be the same.");
        
        // false causes the form submission to be canceled
        return false;
    }
    
    if ( theForm.pw1.value.length < 3 )
    {
        alert( "The password must be at least 3 characters.\n"
                + "Please complete them and Submit again.");
        
        // false causes the form submission to be canceled
        return false;
    }
    
    return true;
}

//-->
</script>
<title>Change Password</title>
</head>

<body onload="doLoad()" bgcolor="#FFFFFF">
<%
DIM	sName

sName = Request("name")
IF 0 < LEN(sName) THEN
%>
  <table border="0" cellspacing="0" cellpadding="4" width="100%">
    <tr>
      <th align="left" bgcolor="#99FFCC">Change Password</th>
      <td align="right" bgcolor="#99FFCC"><a href="teacheredit.asp?name=<%=sName%>">Cancel (Back)</a></td>
    </tr>
  </table>

<form method="POST" action="passwordsubmit.asp" name="pwForm" onsubmit="return checkRequired(this)">
  <input type="hidden" name="name" value="<%=sName%>">
  <table border="0" cellspacing="0" cellpadding="2">
    <tr>
      <td align="right">Enter new password</td>
      <td><input type="password" name="pw1" size="20"></td>
    </tr>
    <tr>
      <td align="right">Confirm new password</td>
      <td><input type="password" name="pw2" size="20"></td>
    </tr>
  </table>
  <p><input type="submit" value="Submit" name="B1"></p>
</form>
<%
END IF
%>
</body>

</html>
<!--#include file="scripts\include_db_end.asp"-->
