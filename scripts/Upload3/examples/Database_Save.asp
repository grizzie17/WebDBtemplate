<!--#INCLUDE FILE="clsUpload.asp"-->
<!--#INCLUDE FILE="Database_Open.asp"-->
<%

' This page is to be called with a query string parameter.
' Example: Database_Download.asp?FileID=367

Dim Sql
Dim FileID
Dim FileName
Dim Upload

Set Upload = New clsUpload

FileName = Upload.Fields("File1").FileName

' Select no records (only schema)
Sql = "Files WHERE FileID = 0"

RecordSet.Open Sql, Connection, 3, 3

' Add a new record
RecordSet.AddNew

RecordSet.Fields("FileName").Value = Upload.Fields("File1").FileName
RecordSet.Fields("FileSize").Value = Upload.Fields("File1").Length
RecordSet.Fields("ContentType").Value = Upload.Fields("File1").ContentType
RecordSet.Fields("BinaryData").AppendChunk Upload("File1").BLOB & ChrB(0)

' Commit the record to the database
RecordSet.Update

' Move to the last record (updates the FileID)
RecordSet.MoveLast

' Get FileID (populated by auto number data type)
FileID = RecordSet.Fields("FileID").Value

Set Upload = Nothing

%>
<!--#INCLUDE FILE="Database_Close.asp"-->

<P>
	File has been saved in database.
</P>
<P>
	View this file:
	<A href="Database_Download.asp?FileID=<%=FileID%>"><%=FileName%></A>
</P>
