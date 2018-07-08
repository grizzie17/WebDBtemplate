<!--#INCLUDE FILE="clsUpload.asp"-->
<%

Dim Upload
Dim FileName
Dim Folder
Dim Ext
Dim FileOK

Set Upload = New clsUpload

' Grab file extension
Ext = Upload.Fields("File1").FileExt

' Check to see if file extension is valid
Select Case Ext
	Case "GIF", "BMP", "PNG", "JPG"
		FileOK = True
	Case Else
		FileOK = False
End Select

' If file was not valid
If Not FileOK Then
	
	' Notify user of error
	Response.Write "Invalid file extension."
	
	' Stop all execution after this line.
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
