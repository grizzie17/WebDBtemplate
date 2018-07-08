<%

DIM	sQuery
sQuery = Request.QueryString

IF "" <> sQuery THEN sQuery = "?" & sQuery

Response.Redirect "../calendar.asp" & sQuery

IF FALSE THEN
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Redirection to Calendar page</title>
</head>

<body>

</body>

</html>
<% 
END IF
%>
