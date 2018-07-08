<%
DIM	sXXName
DIM	sXXAccept
DIM	bXXGood
sXXName = Request("name")
IF 0 = LEN(sXXName) THEN
	Response.Redirect "noname.asp"
END IF
bXXGood = FALSE
sXXAccept = Request.Cookies("accept")
IF "s" = sXXAccept THEN
	bXXGood = TRUE
ELSEIF "u" = sXXAccept  AND  sXXName = Request.Cookies("user") THEN
	bXXGood = TRUE
END IF
IF bXXGood THEN
	Response.Cookies("accept") = sXXAccept
	'Response.Cookies("accept").Expires = DateAdd( "n", 20, DATE() )
ELSE
	DIM	sXXPage
	DIM	sXXID
	
	sXXPage = Request.ServerVariables("URL")
	sXXID = Request("id")
	IF 0 < LEN(sXXID) THEN sXXID = "&id=" & sXXID
	Response.Redirect "pw.asp?name=" & sXXName & "&page=" & sXXPage & sXXID
END IF
%>
