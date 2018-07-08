<%
' Close connections to database
RecordSet.Close
Connection.Close

' Garbage collection
Set RecordSet = Nothing
Set Connection = Nothing
%>