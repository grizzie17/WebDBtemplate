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
    SQLpc = "SELECT * FROM pollcount WHERE ID = (Select MAX(ID) from pollcount)"
    Set RSpc = Conn.Execute(SQLpc)
    SQLpcAdd = "INSERT INTO pollcount (pc)"
    SQLpcAdd = SQLpcAdd & " VALUES ('" & (CInt(RSpc("ID")+1)) & "')"
    Conn.execute(SQLpcAdd) 
    SQLcAdd = "INSERT INTO counted (Total)"
    SQLcAdd = SQLcAdd & " VALUES (0)"
    Conn.execute(SQLcAdd) 
    SQLq = "INSERT INTO ques (q1,q2,q3,q4,q5)"
    SQLq = SQLq & " VALUES ('" & request.form("t1") & "','" & request.form("t2") & "','" & request.form("t3") & "','" & request.form("t4") & "','" & request.form("t5") & "')"
    conn.execute(SQLq) 
    SQLset = "INSERT INTO settings (font,fsize,fcolor,barcolor,bgcolor,tblW)"
    SQLset = SQLset & " VALUES ('" & request.form("D1") & "','" & request.form("D2") & "','" & request.form("D3") & "','" & request.form("D4") & "','" & request.form("D5") & "','" & request.form("D6") & "')"
    conn.execute(SQLset) 
Response.Redirect "index.asp"
%>