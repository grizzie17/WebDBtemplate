<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim Upload
Dim FileName
Dim Folder

Set Upload = New clsUpload

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
