<!--#include file="../../findfiles.asp"-->
<%Response.Buffer=TRUE%>
<%IF Request.Cookies("polllog") <> "xx" THEN Response.Redirect "login.asp"%>
<%
Dim Conn
Dim FilePath
Set Conn = Server.CreateObject("ADODB.Connection")
'FilePath = "C:\Inetpub\wwwroot\poll.mdb" ################################################
'FilePath = Server.MapPath("../poll.mdb")
FilePath = findDatabaseFolder() & "\poll.mdb"
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & FilePath & ";"

SQL = "Update users SET uid = '" & Request.Form("uid") & "', pwd = '" & Request.Form("uid") & "' WHERE ID = 1"
Conn.Execute(SQL)

Conn.Close
SET Conn = NOTHING
Response.Redirect "index.asp"
%>