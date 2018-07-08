<%







CLASS CAnnounceCalendar

	DIM	m_dNowTZ

	DIM	m_sBasePath

	DIM	m_nDate
	DIM	m_nTime
	DIM	m_sTime
	DIM	m_sSubject
	DIM	m_sBody
	DIM m_sLocation
	DIM m_sComments
	DIM	m_sStyle
	DIM	m_sCategories
	DIM	m_sID

	SUB Class_Initialize()
		m_nDate = 0
		m_sTime = ""
		m_sSubject = ""
		m_sBody = ""
		m_sLocation = ""
		m_sComments = ""
		m_sCategories = ""
		m_sID = ""

		m_dNowTZ = DATEADD("h", g_nServerTimeZoneOffset, NOW )
		m_dNowTZ = DATEVALUE(m_dNowTZ)

		
		m_sBasePath = Server.MapPath("announcements")
		
	END SUB
	
	SUB Class_Terminate()
	END SUB
	
	
	PROPERTY LET juliandate( nDate )
		m_nDate = FIX(nDate)
		m_nTime = nDate - m_nDate
	END PROPERTY
	
	PROPERTY LET time( sTime )
		m_sTime = sTime
	END PROPERTY
	
	PROPERTY LET subject( sSubject )
		m_sSubject = sSubject
	END PROPERTY
	
	PROPERTY LET body( sBody )
		m_sBody = sBody
	END PROPERTY
	
	PROPERTY LET location( sLocation )
		m_sLocation = sLocation
	END PROPERTY
	
	PROPERTY LET comments( sComments )
		m_sComments = sComments
	END PROPERTY
	
	PROPERTY LET style( sStyle )
		m_sStyle = sStyle
	END PROPERTY
	
	PROPERTY LET category( sCategory )
		IF "" = m_sCategories THEN
			m_sCategories = sCategory
		ELSE
			m_sCategories = m_sCategories & vbTAB & sCategory
		END IF
	END PROPERTY
	
	PROPERTY LET id( sID )
		m_sID = sID
	END PROPERTY
	
	PROPERTY GET xml
		xml = Nothing
	END PROPERTY
	
	PROPERTY GET xmldom
		SET xmldom = Nothing
	END PROPERTY
	
	SUB outputMessage
	
		IF m_sLocation = "" THEN
			EXIT SUB
		END IF
		
		DIM	oFile
		DIM sPath
		
		sPath = g_oFSO.BuildPath(m_sBasePath, m_sLocation)
		IF g_oFSO.FileExists(sPath) THEN
			SET oFile = g_oFSO.GetFile(sPath)
		ELSE
			EXIT SUB
		END IF
		
		DIM	dDate
		DIM	sSortname
		DIM	sNameLC
		
		dDate = sdateFromJDate( m_nDate )
		'dDate = DATEADD("h", g_nServerTimeZoneOffset, dDate )
		
		
		sNameLC = LCASE(oFile.Name)
		
		IF m_sComments = "" THEN
			sNameLC = LCASE(m_sComments) & sNameLC
		END IF
		
		sSortName = genSortnameFromDate( dDate )
		IF m_dNowTZ <= CDATE(dDate) THEN
			DIM i, j
			i = DATEDIFF( "d", m_dNowTZ, CDATE(dDate) )
			IF i < 7 THEN
				sSortName = STRING( 7-i, "0" ) & sSortName
			ELSE
				j = DATEDIFF( "d", CDATE(dDate ), m_dNowTZ )
				IF j < 8 THEN
					sSortName = genSortnameDeltaFromNow( CDATE(dDate ) )
				END IF
			END IF
		END IF
		sSortname = sSortname & sNameLC


		IF UBOUND(aFileList) <= nFileCount THEN
			REDIM PRESERVE aFileList(nFileCount+20)
		END IF
		aFileList(nFileCount) = sSortName _
				& vbTAB & m_sLocation _
				& vbTAB & oFile.Path _
				& vbTAB & handleSubsText(m_sSubject, m_nDate, m_nTime) _
				& vbTAB & handleSubsText(m_sBody, m_nDate, m_nTime) _
				& vbTAB & "" _
				& vbTAB & oFile.DateLastModified _
				& vbTAB & dDate _
				& vbTAB & dDate _
				& vbTAB & "calendar"
		nFileCount = nFileCount + 1
				
		SET oFile = Nothing

		m_nDate = 0
		m_nTime = 0
		m_sTime = ""
		m_sSubject = ""
		m_sBody = ""
		m_sLocation = ""
		m_sComments = ""
		m_sStyle = ""
		m_sCategories = ""
		m_sID = ""


	END SUB
	
	
	PRIVATE FUNCTION sdateFromJDate( JDate )
	
		DIM	mm
		DIM	dd
		DIM	yy
		
		gregorianFromJDate dd, mm, yy, JDate
		
		sdateFromJDate = CSTR(mm) & "/" & CSTR(dd) & "/" & CSTR(yy)
		
	END FUNCTION
	
	
	PRIVATE FUNCTION handleSubsText( sText, nDate, nTime )
		IF "" <> g_remind_handleSubsText THEN
			handleSubsText = EVAL( g_remind_handleSubsText & "( sText, nDate, nTime )" )
		ELSE
			handleSubsText = sText
		END IF
	END FUNCTION

END CLASS 'CCalendar







SUB appendFileList( nDateBegin, nDateEnd)

	DIM	sRemindFile
	DIM	oCalendar
	DIM	oCalendarFile
	
	DIM	aList()
	REDIM aList(5)
	
	DIM	sRemindFolder
	
	sRemindFolder = findAppDataFolder( "announcements_calendar" )
	IF "" <> sRemindFolder THEN
		DIM oFolder
		SET oFolder = g_oFSO.GetFolder( sRemindFolder )
		IF NOT Nothing IS oFolder THEN
			DIM	nCount
			DIM	oFile
			DIM	sName
			DIM	sNameLC
			DIM	sFile
			DIM	sFileLC
			nCount = 0
			FOR EACH oFile IN oFolder.Files
				sName = oFile.Name
				sNameLC = LCASE(sName)
				IF ".xml" = RIGHT(sNameLC,4) THEN
					sFile = oFile.Path
					IF UBOUND(aList) <= nCount THEN
						REDIM PRESERVE aList(nCount+5)
					END IF
					aList(nCount) = sFile
					nCount = nCount + 1
				END IF
			NEXT 'sFile
			REDIM PRESERVE aList(nCount-1)
		END IF
	END IF
	

	SET	oCalendar = new CAnnounceCalendar
	
	DIM sTemp
	
	
	FOR EACH sTemp IN aList
		IF "" <> sTemp THEN
			SET oCalendarFile = new CCalendarFile
			oCalendarFile.file = sTemp
			oCalendarFile.timeZoneOffset = g_nServerTimeZoneOffset
			oCalendarFile.datebegin = nDateBegin
			oCalendarFile.dateend = nDateEnd
			oCalendarFile.getDates( oCalendar )
		END IF
	NEXT 'sTemp

	IF 0 < nFileCount THEN
		locationSort aFileList, 0, nFileCount-1
		REDIM PRESERVE aFileList(nFileCount-1)
	END IF

	
	SET oCalendar = Nothing
	SET oCalendarFile = Nothing
	
'	FOR EACH sTemp IN aFileList
'		Response.Write Server.HTMLEncode(sTemp) & "<br>"
'	NEXT


END SUB



DIM	dAnnouncementsModified

SUB makeAnnouncements()

	
	
	DIM	nID
	nID = defineTabSortDetails( g_sSiteTabAnnouncements )
	
	DIM	oRS
	SET oRS = getPageListRS( nID )

	IF NOT oRS IS Nothing THEN
	
		IF 0 < oRS.RecordCount THEN

			Response.Write "<div id=""announcecontent"" class=""dynamic"">" & vbCRLF
		
			DIM	oPageRID
			DIM	oFormat
			DIM	oTitle
			DIM	oDescription
			DIM	oPicture
			DIM	oDateModified
			DIM	oCategory
			DIM	oTabID
			DIM	oDateBegin
			DIM	oDateEnd
			DIM	oDateEvent
			
			SET oPageRID = oRS.Fields("RID")
			SET oFormat = oRS.Fields("Format")
			SET oTitle = oRS.Fields("Title")
			SET oDescription = oRS.Fields("Description")
			SET oPicture = oRS.Fields("Picture")
			SET oDateModified = oRS.Fields("DateModified")
			SET oCategory = oRS.Fields("Category")
			SET oTabID = oRS.Fields("TabID")
			SET oDateBegin = oRS.Fields("DateBegin")
			SET oDateEnd = oRS.Fields("DateEnd")
			SET oDateEvent = oRS.Fields("DateEvent")
			
			DIM	nPageRID
			DIM	sFormat
			DIM	sTitle
			DIM	sDescription
			DIM	nPicture
			DIM	dDateModified
			DIM	nCategory
			DIM	nTabID
			DIM	dDateBegin
			DIM	dDateEnd
			DIM	dDateEvent
			
			DIM	sSrcUpdate
			DIM	sAltUpdate
			DIM	sHTMLTitle
			
			DIM	sURL
		
			oRS.MoveFirst
			DO UNTIL oRS.EOF
			
				nPageRID = recNumber(oPageRID)
				sFormat = recString(oFormat)
				sTitle = recString(oTitle)
				sDescription = recString(oDescription)
				nPicture = recNumber(oPicture)
				dDateModified = recDate(oDateModified)
				nCategory = recNumber(oCategory)
				nTabID = recNumber(oTabID)
				dDateBegin = recDate(oDateBegin)
				dDateEnd = recDate(oDateEnd)
				dDateEvent = recDate(oDateEvent)

				IF g_dLocalPageDateModified < dDateModified THEN g_dLocalPageDateModified = dDateModified
				
				sURL = "page.asp?tab=" & nTabID & "&category=" & nCategory & "&page=" & nPageRID
				



				IF 0 < DATEDIFF( "s", g_dLastVisit, dDateModified ) THEN
					sSrcUpdate = "update"
					sAltUpdate = "Updated Announcement/Extra"
				ELSE
					sSrcUpdate = "old"
					sAltUpdate = "Announcement/Extra"
				END IF


				Response.Write "<div class=""BlockHead " & sSrcUpdate & """>" & vbCRLF
				IF dAnnouncementsModified < dDateModified THEN dAnnouncementsModified = dDateModified
				Response.Write "<a href=""" & sURL & """>"
				sHTMLTitle = pagebody_formatString( sFormat, sTitle )
				Response.Write sHTMLTitle
				Response.Write "</a>"
				Response.Write "</div>" & vbCRLF
				Response.Write "<div class=""BlockBody"">" & vbCRLF
				IF 0 < nPicture THEN
					Response.Write "<a href=""" & sURL & """>"
					Response.Write "<img border=""0"" alt=""" & sHTMLTitle & """ src=""picture.asp?id=" & nPicture  & """>" & vbCRLF
					Response.Write "</a>"
				END IF
				IF 0 < LEN(sDescription) THEN
					'Response.Write "<br>"
					Response.Write "<div align=""left"">"
					Response.Write pagebody_formatString( sFormat, sDescription )
					Response.Write "&nbsp;&nbsp; <a href=""" & sURL & """>"
					Response.Write "More..."
					Response.Write "</a>"
					Response.Write "</div>" & vbCRLF
				END IF

				Response.Write "<div style=""text-align: left; color: #999999; font-family: sans-serif; font-size: xx-small;"">"
				Response.Write "Updated: " & Server.HTMLEncode(DATEADD("h", g_nServerTimeZoneOffset, dDateModified))
				Response.Write "</div>"
				Response.Write "</div>"

				oRS.MoveNext
			LOOP
		
		END IF

		Response.Write "</div>" & vbCRLF
	
		oRS.Close
		SET oRS = Nothing
	END IF


END SUB






%>