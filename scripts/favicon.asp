<%

FUNCTION favicon_findPNG( sPath, sName )
	favicon_findPNG = ""

	DIM sFile
	sFile = g_oFSO.BuildPath( sPath, sName & ".png" )
	IF g_oFSO.FileExists( sFile ) THEN
		favicon_findPNG = sName & ".png"
	END IF

END FUNCTION

FUNCTION favicon_findImage( sPath, sName )
	favicon_findImage = ""

	DIM sFile
	sFile = g_oFSO.BuildPath( sPath, sName & ".png" )
	IF g_oFSO.FileExists( sFile ) THEN
		favicon_findImage = sName & ".png"
		EXIT FUNCTION
	END IF

	sFile = g_oFSO.BuildPath( sPath, sName & ".gif" )
	IF g_oFSO.FileExists( sFile ) THEN
		favicon_findImage = sName & ".gif"
		EXIT FUNCTION
	END IF

END FUNCTION


FUNCTION favicon_getText()
	DIM	sData
	sData = ""
	DIM	s
	DIM	i
	DIM	oFolder
	DIM	sIconPath
	DIM sHost
	DIM	sPath
	
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))
	sPath = oFolder.Path
	SET oFolder = Nothing
	sHost = Request.ServerVariables("HTTP_HOST")
	
	s = Request.ServerVariables("URL")
	i = INSTRREV( s, "/" )
	IF 0 < i THEN
		s = LEFT( s, i-1 )
	ELSE
		s = ""
	END IF
	sHost = "http://" & sHost & s & "/"

	sIconPath = g_oFSO.BuildPath( sPath, "favicon.ico" )
	IF g_oFSO.FileExists( sIconPath ) THEN
		sData = sData & "<link rel=""shortcut icon"" type=""image/ico"" href=""http://" & sHost & s & "/favicon.ico"" />" & vbCRLF
	END IF


	DIM	s70, s128
	DIM	s144
	DIM	s150, s270
	DIM	s310x150
	DIM	s310, s588

	s128 = favicon_findImage( sPath, "favicon128" )
	s144 = favicon_findImage( sPath, "favicon144" )
	s150 = favicon_findImage( sPath, "favicon270" )
	s310x150 = favicon_findImage( sPath, "favicon588x270" )
	s310 = favicon_findImage( sPath, "favicon588" )

	s70 = s128
	s150 = s270
	s310 = s588


	' -- for MSWindows 8
	'		image sizes = 128x128, 270x270, 588x270, 588x588
	sData = sData & "<meta name=""msapplication-tooltip"" content=""" & Server.HTMLEncode(g_sSiteName) & """/>" & vbCRLF
	'sData = sData & "<meta name=""msapplication-config"" content=""" & sHost & "srv_msapp_config.asp""/>" & vbCRLF
	IF 0 < (LEN(s70) + LEN(s144) + LEN(s150) + LEN(s310x150) + LEN(s310)) THEN
		IF 0 < LEN(s144) THEN
			sData = sData & "<meta name=""msapplication-TileImage"" content=""" & sHost & s144 & """ />" & vbCRLF
		END IF
		IF 0 < LEN(s70) THEN
			sData = sData & "<meta name=""msapplication-square70x70logo"" content=""" & sHost & s70 & """ />" & vbCRLF
		END IF
		sIconPath = favicon_findImage( sPath, "favicon270" )
		IF 0 < LEN(s150) THEN
			sData = sData & "<meta name=""msapplication-square150x150logo"" content=""" & sHost & s150 & """ />" & vbCRLF
		END IF
		IF 0 < LEN(s310x150) THEN
			sData = sData & "<meta name=""msapplication-wide310x150logo"" content=""" & sHost & s310x150 & """ />" & vbCRLF
		END IF
		IF 0 < LEN(s310) THEN
			sData = sData & "<meta name=""msapplication-square310x310logo"" content=""" & sHost & s310 & """ />" & vbCRLF
		END IF
		IF 0 < (LEN(s150) + LEN(s310x150) + LEN(s310)) THEN
			sData = sData & "<meta name=""msapplication-notification"" " _
						&	"content=""" _
						&	"frequency=360;" _
						&	"polling-uri=" & sHost & "srv_msapp_livetile.asp?q=1;" _
						&	"polling-uri2=" & sHost & "srv_msapp_livetile.asp?q=2;" _
						&	"polling-uri3=" & sHost & "srv_msapp_livetile.asp?q=3;" _
						&	" cycle=1"" />" & vbCRLF
		END IF
	END IF

    ' apple-touch-icon = 57 (png)
    ' 72
    ' 114
    ' 144

	sIconPath = favicon_findPNG( sPath, "apple-touch-icon" )
	IF "" <> sIconPath THEN
		sData = sData & "<link rel=""apple-touch-icon"" href=""" & sHost & sIconPath & """ />" & vbCRLF
	END IF
	if 0 < LEN(s128) THEN
		sData = sData & "<link rel=""apple-touch-icon"" sizes=""128x128"" href=""" & sHost & s128 & """ />" & vbCRLF
	END IF
	if 0 < LEN(s270) THEN
		sData = sData & "<link rel=""apple-touch-icon"" sizes=""270x270"" href=""" & sHost & s270 & """ />" & vbCRLF
	END IF
	if 0 < LEN(s588) THEN
		sData = sData & "<link rel=""apple-touch-icon"" sizes=""588x588"" href=""" & sHost & s588 & """ />" & vbCRLF
	END IF
	' note must be 320 x 480 portrait
	sIconPath = favicon_findPNG( sPath, "apple-touch-startup-320x480" )
	IF "" <> sIconPath THEN
		sData = sData & "<link rel=""apple-touch-startup-image"" href=""" & sHost & sIconPath & """ />" & vbCRLF
	END IF

	favicon_getText = sData
END FUNCTION


SUB favicon_makeLink()

	DIM	sData
	sData = ""
	DIM	oFile
	SET oFile = cache_openTextFile( "site", "favicon.htm", "h", 24, "d" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE

		sData = favicon_getText()
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", "favicon.htm" )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF

	END IF

	IF "" <> sData THEN
		Response.Write sData
	END IF



END SUB
favicon_makeLink

%>