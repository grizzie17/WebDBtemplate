<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim Upload

Set Upload = New clsUpload

' Write debugging information to web page
Response.Write Upload.DebugText

' Release upload object from memory
Set Upload = Nothing

%>
