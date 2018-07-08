<!--#INCLUDE FILE="Database_Open.asp"-->
<%

' This page is to be called with a query string parameter.
' Example: Database_Download.asp?FileID=367

Dim Sql
Dim FileID

' Grab file ID from query string
FileID = Request.QueryString("FileID")

Sql = "SELECT FileName, ContentType, BinaryData FROM Files WHERE FileID = " & FileID

RecordSet.Open Sql, Connection, 3, 3

' If we have at least one record
If Not RecordSet.EOF Then

	' Write out the headers identifying the file
	Response.AddHeader "content-disposition", "inline; filename=" & RecordSet("FileName")
	'Response.AddHeader "content-disposition", "attachment; filename=" & RecordSet("FileName")
	Response.AddHeader "content-length", RecordSet("BinaryData").ActualSize
	Response.ContentType = RecordSet("ContentType")
	
	' Write the file
	Response.BinaryWrite RecordSet("BinaryData")
	
Else

	' Notify the visitor of the problem
	Response.Write("File could not be found")
	
End If

%>

<!--#INCLUDE FILE="Database_Close.asp"-->