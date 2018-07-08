<!--#include file="../findfiles.asp"-->
<%
Dim Conn
Dim FilePath
Set Conn = Server.CreateObject("ADODB.Connection")
'FilePath = "C:\Inetpub\wwwroot\poll.mdb" #############################################
'FilePath = Server.MapPath("poll.mdb")
FilePath = findDatabaseFolder() & "\poll.mdb"
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & FilePath & ";"
    SQLpc = "SELECT * FROM pollcount WHERE ID = (Select MAX(ID) From pollcount)"
    Set RSpc = Conn.Execute(SQLpc)
Response.Cookies("POLL")=RSpc("pc")
Response.Cookies("POLL").Expires="Jan 1, 2010"
    Dim SQL
    Dim SQLa
    SQL = "SELECT * FROM counted WHERE ID = (Select MAX(ID) From pollcount)"
    Set RS = Conn.Execute(SQL)
Dim which
which = Request("R1")
Select Case which
Case 1
x = RS("option1")
y = "option1"
Case 2
x = RS("option2")
y = "option2"
Case 3
x = RS("option3")
y = "option3"
Case 4
x = RS("option4")
y = "option4"
End Select
    Dim intTotal
    Dim intClick
    Dim whichrec
    whichrec = y
    intTotal = RS("total") + 1
    intClick = x + 1
    SQLa = "UPDATE counted SET " & y & " = " & intClick & ", total = "& intTotal &" WHERE ID = " & RSpc("pc")
    Conn.Execute(SQLa)
WhereFrom = Request.ServerVariables("HTTP_REFERER")  
Response.Redirect WhereFrom
%>