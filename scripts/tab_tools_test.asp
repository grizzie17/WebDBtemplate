<%
OPTION EXPLICIT
%>
<!--#include file="tab_tools.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>


<table border="3" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#999999">
	<tr>
		<td bgcolor="#FFFFFF" align="left">


<%

DIM	sTab
sTab = Request.QueryString("tab")

DIM	oTabGen
DIM	oTabData
DIM	oTabFormat

SET oTabGen = NEW CTabGenerate
SET oTabData = NEW CTabDataDummy
SET oTabFormat = NEW CTabFormatDummy

oTabData.URL = "tab_tools_test.asp?tab="

SET oTabGen.TabData = oTabData
SET oTabGen.TabFormat = oTabFormat
oTabGen.Current = sTab
oTabGen.makeTabs


SET oTabFormat = Nothing
SET oTabData = Nothing
SET oTabGen = Nothing


%>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
	<tr>
		<td bgcolor="#FFFFFF">This is some 
		<p>sample text<p>another line<p>yet another line<p>and even another line<p>
		maybe another one<p>yet another line<p>and even another line<p>maybe 
		another one</td>
	</tr>
</table>
		</td>
	</tr>
</table>

</body>

</html>