<%@EnableSessionState=False%>
<!--#INCLUDE FILE="clsProgress.asp"-->
<STYLE>
BODY
{
font-size: 8pt;
}
</STYLE>
<%


' NOTICE:
' If session state is enabled for this page, then 
' data will not be processed until the file has
' been received.

Dim Progress
Dim Percent

' do not cache page
Call Response.AddHeader("pragma","no-cache")
Response.CacheControl = "no-cache"


' Load Progress Information
Set Progress = New clsProgress

' If information could not be loaded
If Not Progress.Load Then
	
	' Garbage collection
	Set Progress = Nothing
	
	' Notify user
	Response.Write "Please Wait ..."
	
	' Instruct browser to refresh
	Refresh()
	
	' Write buffer to browser
	Response.Flush
	
	' Halt execution.
	Response.End
	
End If

With Progress

	Percent = 0

	If Not .TotalBytes = 0 Then
		Percent = Fix((.BytesReceived / .TotalBytes) * 100)
	End If

	Response.Write "Bytes Received: " & .BytesReceived & "<BR>"
	Response.Write "Total Bytes: " & .TotalBytes & "<BR>"
	
	If IsDate(.UploadStarted) Then
		Response.Write "Started: " & FormatDateTime(.UploadStarted, vbLongTime) & "<BR>"
	End If
	
	If IsDate(.LastActive) Then
		Response.Write "Last Active: " & FormatDateTime(.LastActive, vbLongTime) & "<BR>"
	End If
	
	Response.Write "Progress: " & Percent & "%<BR>"
	
	If IsDate(.UploadCompleted) Then
		Response.Write "Completed.<BR>"
	End If
	
	
	' Write progress bar
	%>
	<TABLE border="1" bgColor="red" width="100%" height="16">
	<TR>
	<TD bgcolor="white">
	<SPAN style="height:100%;width:<%=Percent%>%;background-color:blue;"></SPAN>
	</TD>
	<td bgcolor="white" align="center" width="1%"><%=Percent%>%
	</td>
	</TR>
	</TABLE>
	<%

	' If not completed, refresh
	If .UploadCompleted = "" Then
		Refresh
	End If

End With

' Garbage Collection
Set Progress = Nothing

Public Sub Refresh()
		%>
		<SCRIPT>
		window.focus();
		window.setTimeout("window.location.reload()", 1000);
		</SCRIPT>
		<%
End Sub
%>