<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim Upload
Dim FileName
Dim Folder

Set Upload = New clsUpload

' For your reference ...

' 1 KB   = 1024 Bytes
' 10 KB  = 10240 Bytes
' 100 KB = 102400 Bytes

' 1 MB	 = 1048576 Bytes
' 10 MB  = 10485760 Bytes
' 100 MB = 104857600 Bytes

' If file size is greater then 1 MB
If Upload("File1").Length > 1048576 Then

	' Notify user of the error.
	Response.Write "File size must be 1 megabyte or less"
	
	' Stop all execution past this line.
	Response.End
	
End If

' Grab the file name
FileName = Upload.Fields("File1").FileName

' Get path to save file to
Folder = Server.MapPath("Uploads") & "\"

' Save the binary data to the file system
Upload("File1").SaveAs Folder & FileName

' Release upload object from memory
Set Upload = Nothing

%>
<P>
	File has been saved to file system.
</P>
<P>
	View this file:
	<A href="Uploads\<%=FileName%>"><%=FileName%></A>
</P>
