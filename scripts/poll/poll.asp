<!--#include file="../findfiles.asp"-->
<%
Set ConnPoll = Server.CreateObject("ADODB.Connection")
'##################################### USE THE LINE BELOW TO USE A PHYSICAL PATH
'FilePath = "C:\Inetpub\wwwroot\poll.mdb"
'##################################### USE THE LINE BELOW TO USE A VIRTUAL PATH
FilePath = findDatabaseFolder() & "\poll.mdb"
'FilePath = Server.MapPath("poll.mdb")
ConnPoll.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & FilePath & ";"
    SQLpc = "SELECT * FROM pollcount WHERE ID = (Select MAX(ID) From pollcount)"
    Set RSpc = ConnPoll.Execute(SQLpc)

Select Case Request.Cookies("POLL")
Case RSpc("pc")

sql = "Select * from counted WHERE ID = (Select MAX(ID) From counted)"
Set rs1 = ConnPoll.execute(sql)
sqlq = "Select * from ques WHERE ID = (Select MAX(ID) From pollcount)"
Set rsq = ConnPoll.execute(sqlq)
sqlbc = "Select * from settings WHERE ID = (Select MAX(ID) From pollcount)"
Set rsbc = ConnPoll.execute(sqlbc)

pxl = rsbc("tblW")

DIM numArray(3), howWide(3), qArray(3), maxWide, q1, tot
q1 = rsq("q1")
tot = rs1("Total")

qArray(0) = rsq("q2")
qArray(1) = rsq("q3")
qArray(2) = rsq("q4")
qArray(3) = rsq("q5")

Select Case pxl
Case 100
maxWide = 120
Case 125
maxWide = 95
Case 150
maxWide = 75 
Case 175
maxWide = 60 
End Select

numArray(0) = Round((Left(rs1("option1"),5) / rs1("total")) * 100)
numArray(1) = Round((Left(rs1("option2"),5) / rs1("total")) * 100)
numArray(2) = Round((Left(rs1("option3"),5) / rs1("total")) * 100)
numArray(3) = Round((Left(rs1("option4"),5) / rs1("total")) * 100)

FOR i=0 TO UBound(numArray)
IF NOT IsNumeric(numArray(i)) THEN 
numArray(i)=0
END IF
NEXT

Maximum = 0
FOR i=0 TO UBound(numArray)
  IF CInt(numArray(i)) > Maximum THEN
  Maximum = CInt(numArray(i))
  END IF
NEXT

FOR i=0 TO UBound(numArray)
  howWide(i) = FIX((numArray(i)/Maximum)*maxWide)
    IF howWide(i) < 1 THEN
    howWide(i) = 0
    END IF
NEXT

%>
  <table border=0 width="<%=pxl%>" cellspacing="1" cellpadding="1" bgcolor="#<%=rsbc("bgcolor")%>">
  <tr>
  <td colspan="2" align="left"><font face="<%=rsbc("font")%>" size="<%=rsbc("fsize")%>" color="#<%=rsbc("fcolor")%>"><%=q1%></font>&nbsp;</td>
  </tr>
  <%FOR i=0 TO UBound(numarray)
  If qArray(i) <> "" Then
  %>
  <tr>
  <td><font face="<%=rsbc("font")%>" size="<%=rsbc("fsize")%>" color="#<%=rsbc("fcolor")%>"><%=qArray(i)%></font>&nbsp;</td><td align="right"><font face="<%=rsbc("font")%>" size="<%=rsbc("fsize")%>" color="<%=rsbc("fcolor")%>"><%=numArray(i)%> %</font></td>
  </tr>
  <tr>
  <% If numArray(i) > 0 Then %>
  <td colspan=2><img src="img/<%=rsbc("barcolor")%>.jpg" width="<%=Round((numArray(i)/maxWide)*100)%>" height="5"></td>
  <% Else %>
  <td colspan=2><img src="img/xparent.gif" width="50" height="5"></td>
  <% End If %>
  <% End If %>
  </tr>
  <%NEXT%>
  <tr>
  <td colspan="2" align="left"><font face="<%=rsbc("font")%>" size="<%=rsbc("fsize")%>" color="#<%=rsbc("fcolor")%>">Total Votes Cast&nbsp; <%=tot%></font></td>
  </tr>
  </table>
<%
rsq.Close
SET rsq = NOTHING
rsbc.Close
SET rsbc = NOTHING
rs1.Close
SET rs1 = NOTHING

Case Else
dim qArr(4)
sqlq = "Select * from ques WHERE ID = (Select MAX(ID) From pollcount)"
Set rsq = ConnPoll.execute(sqlq)
sqlbc = "Select * from settings WHERE ID = (Select MAX(ID) From pollcount)"
Set rsbc = ConnPoll.execute(sqlbc)
pxl = rsbc("tblW")
q1 = rsq("q1")
qArr(0) = rsq("q2")
qArr(1) = rsq("q3")
qArr(2) = rsq("q4")
qArr(3) = rsq("q5")
%>
<form method="POST" action="update.asp">
<table border="0" width="<%=pxl%>" bgcolor="#<%=rsbc("bgcolor")%>" cellpadding="1" cellspacing="1">
  <tr>
    <td width="100%" colspan="2" align="left"><font face="<%=rsbc("font")%>" size="<%=rsbc("fsize")%>" color="#<%=rsbc("fcolor")%>"><%=q1%></td>
  </tr>
  <%
  x = 1
  For i = 0 TO 3
  If qArr(i) <> "" Then
  %> 
  <tr>
    <td><font face="<%=rsbc("font")%>" size="<%=rsbc("fsize")%>" color="#<%=rsbc("fcolor")%>"><%=qArr(i)%></td>
	<%If i = 0 Then%>
    <td align="right"><input type="radio" value="<%=x%>" checked name="R1"></td>
	<%Else%>
    <td align="right"><input type="radio" value="<%=x%>" name="R1"></td>
	<%End If%>
  </tr>
  <%
  End If
  x = x + 1
  NEXT
  %>
  <tr>
    <td width="100%" colspan="2" align="center"><input type="submit" value="Submit" name="B1"></td>
  </tr>
</table>
</form>
<%
rsq.Close
SET rsq = NOTHING
rsbc.Close
SET rsbc = NOTHING

End Select

RSpc.Close
SET RSpc = NOTHING
ConnPoll.Close
SET ConnPoll = NOTHING
%>