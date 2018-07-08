<%
Dim Connection
Dim RecordSet


' Depending on the drivers installed on your operating system, one method may open
' your database successfully, and another may not.  make adjustments as necessary.

ConnectionString = "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("Files.mdb")
'ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Files.mdb")



' Create object to connect to database
Set Connection = Server.CreateObject("ADODB.Connection")

' Create object to hold data
Set RecordSet = Server.CreateObject("ADODB.Recordset")

' Open database connection
Connection.Open ConnectionString
%>