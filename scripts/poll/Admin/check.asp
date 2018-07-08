<!--#include file="../../findfiles.asp"-->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
'##################################### USE THE LINE BELOW TO USE A PHYSICAL PATH
'FilePath = "C:\Inetpub\wwwroot\poll.mdb"
'##################################### USE THE LINE BELOW TO USE A VIRTUAL PATH
'FilePath = Server.MapPath("../poll.mdb")
FilePath = findDatabaseFolder() & "\poll.mdb"
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & FilePath & ";"

'here we are getting the info from the login form
uid = Request.Form("uid")
pwd = Request.Form("pwd")

'now we will querry the database for a match
SQL = "Select * From users Where uid = '" & uid & "' And pwd = '" & pwd & "'"
Set RS = Conn.Execute(SQL)

'if the user is found we will set the session cookie to xx allowing the user to gain entrance
'i STRONGLY recommend changing "xx" to something unique to your site.......
If Not RS.EOF Then
  Response.Cookies("polllog")="xx"
	
	'always always always destroy recordsets and close connections!
	Set RS = Nothing
	Conn.Close

  'since the user was found, we sent them toodling on to the next page	
  Response.Redirect "index.asp"
Else
   
   'ooops if we got this far they dont know their login info or arent in the database
	'AGAIN always always always destroy recordsets and close connections!
	Set RS = Nothing
	Conn.Close
  'so we send em back to try again	
  Response.Redirect "login.asp"
End If
%>