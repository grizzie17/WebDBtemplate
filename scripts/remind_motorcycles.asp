<%








FUNCTION useMotorcycleFile( sFile, sCategories )
	useMotorcycleFile = FALSE
	DIM	sFileLC
	sFileLC = LCASE(sFile)
	IF 0 < INSTR( sFileLC, "motorcycle" ) THEN
		useMotorcycleFile = TRUE
	ELSEIF 0 < INSTR( sFileLC, "key" ) THEN
		IF "" = sCategories THEN
			useMotorcycleFile = TRUE
		ELSEIF 0 < INSTR(sCategories, "email") THEN
			useMotorcycleFile = TRUE
		ELSEIF 0 < INSTR(sCategories, "key") THEN
			useMotorcycleFile = TRUE
		END IF
	END IF
END FUNCTION


'DIM	dRemindLastModified


FUNCTION loadMotorcycleFiles( nDateBegin, nDateEnd, _
							sCategories, sHolidayCategories, _
							bShowToday, dToday )
	SET loadMotorcycleFiles = loadMotorcycleFiles2( nDateBegin, nDateEnd, _
							sCategories, sHolidayCategories, _
							bShowToday, dToday, "list" )

END FUNCTION





FUNCTION loadMotorcycleFiles2( nDateBegin, nDateEnd, _
							sCategories, sHolidayCategories, _
							bShowToday, dToday, _
							sCalendarView _
							)

	DIM	sRemindFile
	DIM	oCalendar
	DIM	oCalendarFile
	
	DIM	aFileList()
	REDIM aFileList(5)
	DIM	aTempSplit
	DIM	sTemp
	DIM	dateTemp
	
	getRemindList aFileList
	
	
	SET loadMotorcycleFiles2 = Nothing
	
	IF -1 < UBOUND(aFileList) THEN
		aTempSplit = SPLIT( aFileList(0), vbTAB )
		dRemindLastModified = CDATE(aTempSplit(2))
		
		SET	oCalendar = new CCalendar
		
		IF bShowToday THEN
			oCalendar.juliandate = jdateFromVBDate( dToday )
			oCalendar.subject = "* T * O * D * A * Y *"
			oCalendar.style = "RmdToday"
			oCalendar.time = "00:00"
			oCalendar.outputMessage
		END IF
		
		DIM oXML

		FOR EACH sTemp IN aFileList
			IF useMotorcycleFile( sTemp, sCategories ) THEN
				IF "" <> sTemp THEN
					aTempSplit = SPLIT( sTemp, vbTAB )
					SET oXML = remindLoadWithInclude( aTempSplit(1) )
					IF NOT oXML IS Nothing THEN
						dateTemp = CDATE( aTempSplit(2) )
						IF dRemindLastModified < dateTemp THEN dRemindLastModified = dateTemp
						SET oCalendarFile = new CCalendarFile
						SET oCalendarFile.xmldom = oXML
						'oCalendarFile.file = aTempSplit(1)
						oCalendarFile.datebegin = nDateBegin
						oCalendarFile.dateend = nDateEnd
						oCalendarFile.calendarView = sCalendarView
						IF 0 < INSTR(aTempSplit(1),";categories") THEN
							IF 0 < LEN(sCategories) THEN oCalendarFile.categoryQuery = sCategories
						ELSEIF 0 < INSTR(aTempSplit(1),";holiday") THEN
							IF 0 < LEN(sHolidayCategories) THEN oCalendarFile.categoryQuery = sHolidayCategories
						ELSE
							IF 0 < LEN(sCategories) THEN oCalendarFile.categoryQuery = sCategories
						END IF
						oCalendarFile.getDates( oCalendar )
					END IF
				END IF
			END IF
		NEXT 'sTemp
				
		SET loadMotorcycleFiles2 = oCalendar
		SET oCalendar = Nothing
		SET oCalendarFile = Nothing

	END IF
END FUNCTION



%>