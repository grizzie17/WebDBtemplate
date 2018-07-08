<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="robots" content="none">
</head>

<body>

<%

IF "" <> Request.Form THEN

	IF "1701" = Request.Form("code") THEN
	
		DIM	bMore
		bMore = FALSE

		DO
			bMore = FALSE
			FOR EACH sKey IN Application.Contents
				IF "RCW" = LEFT(sKey,3) THEN
					bMore = TRUE
					Application.Contents.Remove( sKey )
				END IF
			NEXT 'sKey
		LOOP WHILE bMore
		
		DO
			bMore = FALSE
			FOR EACH sKey IN Session.Contents
				IF "RCW" = LEFT(sKey,3) THEN
					bMore = TRUE
					Session.Contents.Remove( sKey )
				END IF
			NEXT 'sKey
		LOOP WHILE bMore
	
	END IF

END IF



DIM sCookie
DIM	sKey

Response.Write "<h2>Application</h2>"

FOR EACH sKey IN Application.Contents
	IF "RCW" = LEFT(sKey,3) THEN
		sCookie = Application.Contents( sKey )
		Response.Write "<p>" & sKey & " = " & Server.HTMLEncode( Replace( sCookie, ",", ", " ) ) & "</p>" & vbCRLF
	END IF
NEXT 'sKey

Response.Write "<h2>Session</h2>"

FOR EACH sKey IN Session.Contents
	IF "RCW" = LEFT(sKey,3) THEN
		sCookie = Session.Contents( sKey )
		Response.Write "<p>" & sKey & " = " & Server.HTMLEncode( Replace( sCookie, ",", ", " ) ) & "</p>" & vbCRLF
	END IF
NEXT 'sKey


%>

<form method="POST" action="cookie.asp">
	<p><input type="text" name="code" size="20"><br>
	<input type="hidden" name="RCW_gallery" value="">
	<input type="hidden" name="RCWhumor" value="">
	<input type="hidden" name="RCWsafety" value="">
	<input type="submit" value="Reset Cookies" name="B1"></p>
</form>

</body>

</html>