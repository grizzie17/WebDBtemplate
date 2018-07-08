<!--#INCLUDE FILE="clsUpload.asp"-->
<!--#INCLUDE FILE="clsImage.asp"-->
<%

Dim Upload
Dim FileName
Dim Folder
Dim Image
Dim FileOK
Dim Width
Dim Height

Set Upload = New clsUpload

Set Image = New clsImage

Image.DataStream = Upload("File1").BLOB()

Width = Image.Width
Height = Image.Height

Set Image = Nothing

' Initialize file status to OK
FileOK = True

If Width > 640 Then FileOK = False
If Height > 480 Then FileOK = False

' If the image didn't fall within the requirements
If Not FileOK Then
	
	' Notify the user
	Response.Write "Image size must be 640x480.  Your image was " & Width & "x" & Height
	
	' Stop all execution after this line
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
<IMG src="Uploads\<%=FileName%>" width="<%=Width%>" height="<%=Height%>">
