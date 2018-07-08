<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim Upload
Dim FileName1
Dim FileName2
Dim Folder

Set Upload = New clsUpload

' Grab the file names
FileName1 = Upload.Fields("File1").FileName
FileName2 = Upload.Fields("File2").FileName

' Get path to save file to
Folder = Server.MapPath("Uploads") & "\"

' Save the binary data to the file system
Upload("File1").SaveAs Folder & FileName1
Upload("File2").SaveAs Folder & FileName2

' Release upload object from memory
Set Upload = Nothing

%>
<P>
	Files have been saved to file system.
</P>
<P>
	View these files:<BR>
	<A href="Uploads\<%=FileName1%>"><%=FileName1%></A><BR>
	<A href="Uploads\<%=FileName2%>"><%=FileName2%></A><BR>
</P>
