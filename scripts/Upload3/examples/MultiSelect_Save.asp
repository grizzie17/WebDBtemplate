<!--#INCLUDE FILE="clsUpload.asp"-->
<%
Dim Upload
Dim FileName
Dim Folder
Dim Colors
Dim Color
Dim ColorList

Set Upload = New clsUpload

' Grab the file name
FileName = Upload.Fields("File1").FileName

' Get path to save file to
Folder = Server.MapPath("Uploads") & "\"

' Save the binary data to the file system
Upload("File1").SaveAs Folder & FileName

' Grab all selected colors
Colors = Upload.Collection("Colors")

' Loop through each field
For Each Color In Colors
	
	' build list of colors
	ColorList = ColorList & Color.Value & ", "

Next

' Release upload object from memory
Set Upload = Nothing
%>
<P>
	Your favorite colors appear to be: <%=ColorList%>
</P>
<P>
	File has been saved to file system.
</P>
<P>
	View this file:
	<A href="Uploads\<%=FileName%>"><%=FileName%></A>
</P>