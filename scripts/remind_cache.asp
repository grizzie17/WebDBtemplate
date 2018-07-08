<%


FUNCTION remind_cache( sFile, sInterval, nIntValue, sBreakInterval, _
					nDateBegin, nDateEnd, sCategories, sHolidayCategories, bShowToday, dToday, _
					sXSLTFile )

	gHtmlOption_encodeEmailAddresses = TRUE
	
	DIM	sHtmData
	sHtmData = ""
	DIM	oFile
	SET oFile = cache_openTextFile( "remind", sFile, sInterval, nIntValue, sBreakInterval )
	IF NOT oFile IS Nothing THEN
		sHtmData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
		SET oFile = cache_openTextFile( "remind", "RemindLastModified.txt", sInterval, nIntValue, sBreakInterval )
		IF NOT oFile IS Nothing THEN
			DIM	sDate
			sDate = oFile.ReadLine
			oFile.Close
			SET oFile = Nothing
			IF "" <> sDate THEN
				dRemindLastModified = CDATE(sDate)
			END IF
		END IF
	ELSE
		DIM	oCalendar
		SET oCalendar = loadRemindFiles( nDateBegin, nDateEnd, sCategories, sHolidayCategories, bShowToday, dToday )
		IF NOT oCalendar IS Nothing THEN
			DIM	oXML
			SET oXML = oCalendar.xmldom
			SET oCalendar = Nothing
			
			DIM	oXSL
			SET oXSL = remindLoadXmlFile( sXSLTFile )
			sHtmData = TRIM(oXML.transformNode(oXSL))
			SET oXML = Nothing
			SET oXSL = Nothing
			
			SET oFile = cache_makeFile( "remind", sFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sHtmData
				oFile.Close
				SET oFile = Nothing
			END IF
			SET oFile = cache_makeFile( "remind", "RemindLastModified.txt" )
			IF NOT oFile IS Nothing THEN
				oFile.Write CSTR(dRemindLastModified) & vbCRLF
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF


	
	remind_cache = sHtmData

END FUNCTION



%>