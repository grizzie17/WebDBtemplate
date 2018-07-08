<%
' these routines use include_cache.asp

FUNCTION remind_buildCssList()
	remind_buildCssList = ""

	DIM	oFile
	DIM	sFile
	DIM	sText
	DIM	sTemp
	DIM	sClassname
	DIM	sDesc
	
	sTemp = "<ul class=""remindcsslist"">"
	
	sFile = findRemindFile( "remind.css" )
	SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
	sText = oFile.ReadAll()
	oFile.Close
	SET oFile = Nothing
	
	DIM	oReg
	DIM	oMatchList
	DIM	oMatch
	
	SET oReg = NEW RegExp
	oReg.Pattern = "\s\.(\w+)\s*\{\s*/\*\s*::\s*([^:]+)(::|\*/)"
	oReg.IgnoreCase = TRUE
	oReg.Global = TRUE
	SET oMatchList = oReg.Execute( sText )
	FOR EACH oMatch IN oMatchList
		sClassname = oMatch.Submatches(0)
		sDesc = TRIM(oMatch.Submatches(1))
		IF "" <> sDesc THEN
			IF "RmdToday" <> sClassname THEN
				sTemp = sTemp & "<li class=""" & sClassname & """>"
				sTemp = sTemp & Server.HTMLEncode(sDesc)
				sTemp = sTemp & "</li>" & vbCRLF
			END IF
		END IF
	NEXT 'oMatch
	sTemp = sTemp & "</ul>" & vbCRLF

	remind_buildCssList = sTemp

END FUNCTION

SUB remind_liststyles()

	DIM	sList
	sList = ""
	DIM	oFile
	SET oFile = cache_OpenTextFile( "remind", "csslist.htm", "d", 28, "m" )
	IF NOT oFile IS Nothing THEN
		sList = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sList = remind_buildCssList()
		IF "" <> sList THEN
			SET oFile = cache_makeFile( "remind", "csslist.htm" )
			IF NOT oFile IS Nothing THEN
				oFile.Write sList
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
	IF "" <> sList THEN
		Response.Write sList
	END IF

END SUB



%>