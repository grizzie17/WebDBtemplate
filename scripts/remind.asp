<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2008 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------
%>
<!--#include file="dateutils.asp"-->
<!--#include file="htmlformat.asp"-->
<%

CONST kTropicalYear = 365.242190


DIM g_remind_loadXmlFile
g_remind_loadXmlFile = "remind_loadXmlFileWithXinclude"

DIM	g_remind_handleSubsText
g_remind_handleSubsText = "remind_handleSubsText"

g_htmlFormat_metaSubsFunc = "remind_metaSubsFunc"
g_htmlFormat_pictureFunc = "remind_pictureFile"

	
	
SUB appendList( n, a, v )
	IF UBOUND(a) < n+1 THEN REDIM PRESERVE a(n+10)
	a(n) = v
	n = n + 1
END SUB



DIM	g_sPagePicturePath
DIM	g_sPagePictureVirtPath
g_sPagePicturePath = ""
g_sPagePictureVirtPath = ""

' this function depends on the global "sName" to define the faculty name-id
FUNCTION remind_pictureFile( sLabel )
	DIM	sFile
	IF 0 = LEN(g_sPagePicturePath) THEN
		DIM	sProfilePath
		DIM	sPicturePath

		sPicturePath = findFolder( "images\remind" )
		IF 0 < LEN(sPicturePath) THEN
			DIM	sRootPath
			g_sPagePicturePath = sPicturePath
			sRootPath = Server.MapPath("/")
			sPicturePath = MID(g_sPagePicturePath,LEN(sRootPath)+1)
			g_sPagePictureVirtPath = REPLACE(sPicturePath,"\","/",1,-1,vbTextCompare) & "/"
		END IF
	END IF
	sFile = sLabel & ".gif"
	IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( g_sPagePicturePath, sFile )) THEN
		sFile = sLabel & ".jpg"
		IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( g_sPagePicturePath, sFile )) THEN
			sFile = sLabel & ".png"
			IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( g_sPagePicturePath, sFile )) THEN sFile = ""
		END IF
	END IF
	IF 0 < LEN(sFile) THEN
			sFile = g_sPagePictureVirtPath & sFile
	END IF
	remind_pictureFile = sFile
END FUNCTION




FUNCTION f_year( aArgs(), nDate, nTime )
		IF 0 < UBOUND(aArgs) THEN
			IF ISNUMERIC( aArgs(1) ) THEN
				DIM	yy
				DIM	d, m, y
				DIM	s, t
				
				yy = CINT(aArgs(1))
				gregorianFromJDate d, m, y, nDate
				d = y - yy
				s = CSTR(d)
				t = RIGHT(s,2)
				IF 1 < LEN(s)  AND  "1" = LEFT(t,1) THEN
					s = s & "th"
				ELSE
					SELECT CASE RIGHT(s,1)
					CASE "1"
						s = s & "st"
					CASE "2"
						s = s & "nd"
					CASE "3"
						s = s & "rd"
					CASE ELSE
						s = s & "th"
					END SELECT
				END IF
				f_year = s
			ELSE
				f_year = "{{invalid year specified}}"
			END IF
		ELSE
			f_year = "{{year not specified}}"
		END IF
END FUNCTION
	
FUNCTION f_days( aArgs(), nDate, nTime )
		DIM	n
		
		n = nDate - jdateFromVBDate( NOW )
		f_days = CSTR(n)
END FUNCTION
	
FUNCTION f_time( aArgs(), nDate, nTime )
		DIM	hh, mm, ss
		DIM	n
		DIM	sTemp

		n = FIX(nTime * CDBL(3600*24))
		hh = INT( n / 3600 )
		n = n - (hh*3600)
		mm = INT( n / 60 )

		sTemp = ""
		IF 0 < hh THEN
			IF hh < 10 THEN
				sTemp = "0" & hh & ":"
			ELSE
				sTemp = "" & hh & ":"
			END IF
		ELSE
			sTemp = "00:"
		END IF
		IF 0 < mm THEN
			IF mm < 10 THEN
				sTemp = sTemp & "0" & mm
			ELSE
				sTemp = sTemp & mm
			END IF
		ELSE
			sTemp = sTemp & "00"
		END IF
		
		f_time = sTemp
END FUNCTION


FUNCTION f_datetime( aArgs(), nDate, nTime )
	DIM	mm
	DIM	dd
	DIM	yy
	DIM	sMonName
	DIM sWeekday
	gregorianFromJDate dd, mm, yy, nDate
	IF 0 < UBOUND(aArgs) THEN
		SELECT CASE LCASE(aArgs(1))
		CASE "tiny", "short", "noyear", "mm/dd"
			f_datetime = CSTR(mm) & "/" & CSTR(dd)
		CASE "general", "normal", "mm/dd/yy"
			f_datetime = CSTR(mm) & "/" & CSTR(dd) & "/" & CSTR(yy)
		CASE "long"
			sMonName = MONTHNAME(mm)
			f_datetime = sMonName & " " & dd & ", " & yy
		CASE "verbose"
			sMonName = MONTHNAME(mm)
			sWeekday = WEEKDAYNAME( weekdayFromJDate(nDate), FALSE, 1 )
			f_datetime = sWeekday & ", " & sMonName & " " & dd & ", " & yy
		END SELECT
	ELSE
		f_datetime = CSTR(mm) & "/" & CSTR(dd) & "/" & CSTR(yy)
	END IF
END FUNCTION
	
FUNCTION f_weekdays( aArgs(), nDate, nTime )
		DIM	n
		DIM	bInclusive
		DIM	jdNOW
		IF 0 < UBOUND(aArgs) THEN
			SELECT CASE LCASE(aArgs(1))
			CASE "include", "inclusive", "inclussive"
				bInclusive = TRUE
			CASE ELSE
				bInclusive = FALSE
			END SELECT
		ELSE
			bInclusive = FALSE
		END IF
		jdNOW = jdateFromVBDate( NOW() )
		n = diffWeekdaysJDates( jdNOW, nDate, bInclusive )
		f_weekdays = CSTR(n)
END FUNCTION
	

DIM	g_remind_nDate
DIM	g_remind_nTime

FUNCTION remind_metaSubsFunc( aArgs, sText )
	DIM	nDate
	DIM	nTime
	
	IF "" = sText THEN
		nDate = g_remind_nDate
		nTime = g_remind_nTime
		SELECT CASE LCASE( aArgs(0) )
		CASE "year", "years", "anniv"
			remind_metaSubsFunc = f_year( aArgs, nDate, nTime )
		CASE "day", "days"
			remind_metaSubsFunc = f_days( aArgs, nDate, nTime )
		CASE "weekdays", "weekday", "workdays", "workdays"
			remind_metaSubsFunc = f_weekdays( aArgs, nDate, nTime )
		CASE "time"
			remind_metaSubsFunc = f_time( aArgs, nDate, nTime )
		CASE "date", "datetime"
			remind_metaSubsFunc = f_datetime( aArgs, nDate, nTime )
		CASE ELSE
			remind_metaSubsFunc = ""
		END SELECT
	ELSE
		remind_metaSubsFunc = ""
	END IF
END FUNCTION

	
FUNCTION remind_handleSubsText( sText, nDate, nTime )
	DIM	sTemp
	g_remind_nDate = nDate
	g_remind_nTime = nTime
	sTemp = htmlFormatEncodeText( sText )
	IF 0 < INSTR(1,sText,"{{",vbTextCompare) THEN sTemp = htmlFormatMeta( sTemp )
	remind_handleSubsText = sTemp
END FUNCTION






'==== CCalendar ========================= CCalendar =====








CLASS CCalendar

	DIM	m_oXML
	DIM	m_oRoot
	
	DIM	m_nDate
	DIM	m_nTime
	DIM	m_sTime
	DIM	m_sSubject
	DIM	m_sContent
	DIM m_sLocation
	DIM m_sComments
	DIM	m_sStyle
	DIM	m_sCategories
	DIM	m_sID

	SUB Class_Initialize()
		m_nDate = 0
		m_sTime = ""
		m_sSubject = ""
		m_sContent = ""
		m_sLocation = ""
		m_sComments = ""
		m_sCategories = ""
		m_sID = ""

		SET m_oXML = Server.CreateObject("Msxml2.DOMDocument.6.0")
		m_oXML.async = false
		m_oXML.setProperty "ResolveExternals", FALSE

		DIM	oXRoot
		SET oXRoot = m_oXML.createProcessingInstruction( "xml", "version=""1.0"" encoding=""ISO-8859-1""")
		m_oXML.appendChild( oXRoot )
		SET oXRoot = Nothing

	
		'SET m_oXML = Server.CreateObject("msxml2.DOMDocument")
		'SET m_oXML = Server.CreateObject("Microsoft.XMLDOM")
		
		' Create XML prolog
		SET m_oRoot = m_oXML.createNode( 1, "calendar", "" )
		m_oXML.appendChild( m_oRoot )
	END SUB
	
	SUB Class_Terminate()
		SET m_oRoot = Nothing
		SET m_oXML = Nothing
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
	
	PROPERTY LET content( sContent )
		m_sContent = sContent
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
		xml = m_oXML.xml
	END PROPERTY
	
	PROPERTY GET xmldom
		SET xmldom = m_oXML
	END PROPERTY
	
	SUB outputMessage
		DIM	oEvent
		DIM	oDate
		DIM	oStyle
		DIM	oSubject
		DIM	oContent
		DIM	oText
		DIM	oAttr
		
		IF "" = m_sTime THEN
			m_sTime = "99:99"
		END IF
		
		SET oEvent = m_oXML.createNode( 1, "event", "" )
		SET oText = m_oXML.createTextNode( logDateFromJDate( m_nDate ) )
		SET oDate = m_oXML.createNode( 1, "date", "" )
		
		DIM	dd, mm, yy
		gregorianFromJDate dd, mm, yy, m_nDate
		oDate.setAttribute "d", dd
		oDate.setAttribute "m", mm
		'oDate.setAttribute "mname", monthname(mm,3)
		oDate.setAttribute "y", yy
		oDate.setAttribute "wd", weekdayFromJDate( m_nDate )
		IF 0 < LEN(m_sTime) THEN
			oDate.setAttribute "time", m_sTime
		END IF

		oDate.appendChild( oText )
		oEvent.appendChild( oDate )
		
		IF 0 < LEN( m_sStyle ) THEN
			SET oText = m_oXML.createTextNode( m_sStyle )
			SET oStyle = m_oXML.createNode( 1, "style", "" )
			oStyle.appendChild( oText )
			oEvent.appendChild( oStyle )
		END IF
		SET oText = m_oXML.createCDATASection( handleSubsText(m_sSubject, m_nDate, m_nTime) )
		SET oSubject = m_oXML.createNode( 1, "subject", "" )
		oSubject.appendChild( oText )
		oEvent.appendChild( oSubject )
		
		IF 0 < LEN( m_sContent ) THEN
			SET oText = m_oXML.createCDATASection( handleSubsText(m_sContent, m_nDate, m_nTime) )
			SET oContent = m_oXML.createNode( 1, "content", "" )
			oContent.appendChild( oText )
			oEvent.appendChild( oContent )
		END IF
		
		IF 0 < LEN( m_sLocation ) THEN
			SET oText = m_oXML.createCDATASection( handleSubsText(m_sLocation, m_nDate, m_nTime) )
			SET oContent = m_oXML.createNode( 1, "location", "" )
			oContent.appendChild( oText )
			oEvent.appendChild( oContent )
		END IF
		
		IF 0 < LEN( m_sComments ) THEN
			SET oText = m_oXML.createCDATASection( handleSubsText(m_sComments, m_nDate, m_nTime) )
			SET oContent = m_oXML.createNode( 1, "comments", "" )
			oContent.appendChild( oText )
			oEvent.appendChild( oContent )
		END IF
		
		IF 0 < LEN( m_sCategories ) THEN
			DIM	aSplitCat
			DIM	oCat
			DIM	sCat
			aSplitCat = SPLIT(m_sCategories,vbTAB)
			FOR EACH sCat IN aSplitCat
				SET oText = m_oXML.createTextNode( sCat )
				SET oCat = m_oXML.createNode( 1, "category", "" )
				oCat.appendChild( oText )
				oEvent.appendChild( oCat )
			NEXT 'sCat
			SET oCat = Nothing
			SET oText = Nothing
		END IF
		m_oRoot.appendChild( oEvent )
		
		SET oEvent = Nothing
		SET oDate = Nothing
		SET oSubject = Nothing
		SET oContent = Nothing
		SET oText = Nothing
		m_nDate = 0
		m_nTime = 0
		m_sTime = ""
		m_sSubject = ""
		m_sContent = ""
		m_sLocation = ""
		m_sComments = ""
		m_sStyle = ""
		m_sCategories = ""
		m_sID = ""
	END SUB
	
	
	PRIVATE FUNCTION handleSubsText( sText, nDate, nTime )
		IF "" <> g_remind_handleSubsText THEN
			handleSubsText = EVAL( g_remind_handleSubsText & "( sText, nDate, nTime )" )
		ELSE
			handleSubsText = sText
		END IF
	END FUNCTION

END CLASS 'CCalendar






'==== CCalendarFile ========================= CCalendarFile =====



FUNCTION remind_loadXmlFile( sFile )

	SET remind_loadXmlFile = Nothing

	DIM	oXML
	SET oXML = Server.CreateObject("Msxml2.DOMDocument.6.0")
	oXML.async = false
	oXML.setProperty "ProhibitDTD", false
	oXML.setProperty "ResolveExternals", true
	oXML.setProperty "AllowDocumentFunction", true
	'oXML.setProperty "AllowXmlAttributes", true
	'oXML.setProperty "ValidateOnParse", true
	'oXML.validateOnParse = false
	oXML.load( sFile )

	IF oXML.parseError.errorCode <> 0 THEN
		Response.Write "Error in File: " & Server.HTMLEncode(sFile) & "<br>" & vbCRLF
		Response.Write "Error Code: " & oXML.parseError.errorCode & "<br>" & vbCRLF
		Response.Write "Error Reason: " & oXML.parseError.reason & "<br>" & vbCRLF
		Response.Write "Error Line: " & oXML.parseError.line & "<br>" & vbCRLF
	ELSE
		SET remind_loadXmlFile = oXML
	END IF
	
	SET oXML = Nothing

END FUNCTION


FUNCTION remind_loadXmlFileWithXinclude( sFile )

	SET remind_loadXmlFileWithXinclude = Nothing
	
	DIM	oXML
	SET oXML = remind_loadXmlFile( sFile )
	IF NOT oXML IS Nothing THEN

		DIM	oInc
		SET oInc = Nothing
		ON ERROR Resume Next
		SET oInc = oXml.selectSingleNode("/reminders/event")
		ON ERROR Goto 0
		IF NOT oInc IS Nothing THEN
		
			SET oInc = Nothing
			SET remind_loadXmlFileWithXinclude = oXml
		
		ELSE
		
		
			DIM oXSL
			SET oXSL = Server.CreateObject("Msxml2.DOMDocument.6.0")
			oXSL.async = false
			oXSL.setProperty "ProhibitDTD", false
			oXSL.setProperty "ResolveExternals", true
			oXSL.setProperty "AllowDocumentFunction", true
			DIM	sXSLFile
			sXSLFile = Server.MapPath("scripts/remind_copy.xslt")
			oXSL.load( sXSLFile )
			IF oXSL.parseError.errorCode <> 0 THEN
				Response.Write "Error in File: " & Server.HTMLEncode(sXSLFile) & "<br>" & vbCRLF
				Response.Write "Error Code: " & oXSL.parseError.errorCode & "<br>" & vbCRLF
				Response.Write "Error Reason: " & oXSL.parseError.reason & "<br>" & vbCRLF
				Response.Write "Error Line: " & oXSL.parseError.line & "<br>" & vbCRLF
			END IF



			DIM oDstXml
			DIM	xStr
			xStr = oXML.transformNode(oXSL)
			
			SET oDstXml = Server.CreateObject("Msxml2.DOMDocument.6.0")
			oDstXml.async = false
			oDstXml.loadxml( xStr )
			
			SET remind_loadXmlFileWithXinclude = oDstXml
			
			SET oDstXml = Nothing
			SET oXSL = Nothing
		
		END IF
		
		SET oXml = Nothing

	END IF

END FUNCTION








CLASS CCalendarFile

	PRIVATE m_oXML
	PRIVATE m_sFilename
	
	PRIVATE m_sCalendarView
	PRIVATE	m_sCategories
	PRIVATE	m_sCategoryFilter
	PRIVATE	m_nDateBegin	'julian
	PRIVATE	m_nDateEnd		'julian
	PRIVATE	m_nPending		'number of days before now
	PRIVATE m_nTimeZoneOffset
	
	PRIVATE	m_nYearBegin
	PRIVATE	m_nYearNow
	
	SUB Class_Initialize()
		'SET m_oXML = Server.CreateObject("msxml2.DOMDocument")
		'SET m_oXML = Server.CreateObject("Msxml2.DOMDocument.6.0")
		'SET m_oXML = Server.CreateObject("Microsoft.XMLDOM")
		SET m_oXML = Nothing
		m_sFilename = ""
		
		m_sCalendarView = ""	'process all durations
		m_sCategories = ""	'all
		m_sCategoryFilter = ""
		m_nDateBegin = 0	'all
		m_nDateEnd = 999999999
		m_nPending = 0
		m_nTimeZoneOffset = 0
		
		m_nYearBegin = 0
		m_nYearNow = Year( Now )
	END SUB
	
	SUB Class_Terminate()
		SET m_oXML = Nothing
	END SUB


	PRIVATE FUNCTION loadXmlFile( sFile )

		SET loadXmlFile = EVAL( g_remind_LoadXmlFile & "( sFile )" )

	END FUNCTION


	
	PROPERTY LET file( sFilename )
		m_sFilename = sFilename
		
		SET m_oXML = loadXmlFile( m_sFilename )
	END PROPERTY
	
	
	PROPERTY SET xmldom( oXmlDocument )
		SET m_oXML = oXmlDocument
	END PROPERTY
	
	
	PROPERTY LET calendarView( sView )
		SELECT CASE LCASE(sView)
		CASE "list"
			m_sCalendarView = "list"
		CASE "month-at-a-glance", "mag"
			m_sCalendarView = "mag"
		CASE ELSE
			m_sCalendarView = ""
		END SELECT
	END PROPERTY
	
	
	PROPERTY LET timeZoneOffset( nTZ )
		m_nTimeZoneOffset = nTZ
	END PROPERTY
	
	PROPERTY LET datebegin( nDateBegin )
		DIM	dd, mm, yy
		IF 0 < nDateBegin THEN
			m_nDateBegin = nDateBegin
			gregorianFromJDate dd, mm, yy, nDateBegin
			m_nYearBegin = yy
		ELSE
			m_nDateBegin = 0
			m_nYearBegin = 0
		END IF
	END PROPERTY
	
	PROPERTY LET dateend( nDateEnd )
		DIM	n
		n = CLNG(nDateEnd)
		IF 0 < n THEN
			m_nDateEnd = n
		ELSE
			pending = ABS( n )
		END IF
	END PROPERTY
	
	PROPERTY LET pending( nDays )
		m_nPending = ABS(CLNG(nDays))
		'm_nDateEnd = jdateFromVBDate( NOW )
		'm_nDateBegin = m_nDateEnd - m_nPending
		DIM	dNow
		dNow = DATEADD( "h", m_nTimeZoneOffset, NOW )
		m_nDateBegin = jdateFromVBDate( dNow )
		m_nDateEnd = m_nDateBegin + m_nPending - 1
	END PROPERTY
	
	PRIVATE SUB buildCategoriesFilter( sCategories, oper )
		DIM	aSplitDate
		IF 0 < LEN(sCategories) THEN
			m_sCategories = sCategories
			aSplitDate = SPLIT( sCategories, ",", -1, vbTextCompare )
			DIM	sTemp, sBuild
			IF UBOUND(aSplitDate) < 1 THEN
				IF sCategories = "none" THEN
					m_sCategoryFilter = "[ not(category) ]"
				ELSE
					m_sCategoryFilter = "[ category = """ & sCategories & """]"
					'm_sCategoryFilter = "[ $any$ category = """ & sCategories & """]"
				END IF
			ELSE
				sBuild = ""
				FOR EACH sTemp IN aSplitDate
					IF sTemp = "none" THEN
						IF 0 < LEN(sBuild) THEN
							sBuild = sBuild & " " & oper & " not(category)"
						ELSE
							sBuild = "not(category)"
						END IF
					ELSE
						IF 0 < LEN(sBuild) THEN
							sBuild = sBuild & " " & oper & " ( category = """ & sTemp & """)"
							'sBuild = sBuild & " " & oper & " ( $any$ category = """ & sTemp & """)"
						ELSE
							sBuild = "($any$ category = """ & sTemp & """)"
							'sBuild = "( category = """ & sTemp & """)"
						END IF
					END IF
				NEXT 'sTemp
				m_sCategoryFilter = "[ " & sBuild & " ]"
			END IF
			'Response.Write "category-filter: " & m_sCategoryFilter
		ELSE
			m_sCategories = ""
			m_sCategoryFilter = ""
		END IF
	END SUB
	
	PROPERTY LET categories( sCategories )
		buildCategoriesFilter sCategories, "or"
	END PROPERTY
	PROPERTY LET categoriesUnion( sCategories )
		buildCategoriesFilter sCategories, "or"
	END PROPERTY
	PROPERTY LET categoriesIntersect( sCategories )
		buildCategoriesFilter sCategories, "and"
	END PROPERTY
	
	PROPERTY LET categoryQuery( sQuery )
		' comma list is 'or'
		' colon list is 'and'
		' tilde is 'not'
		' no grouping
		
		IF "" <> sQuery THEN

			m_sCategories = sQuery
			m_sCategoryFilter = ""
		
			DIM	sBuildMega
			DIM	sBuild
			DIM	sF
			DIM	aOrs
			DIM	sOrItem
			DIM	aAnds
			DIM	sAndItem
			sBuildMega = ""
			sBuild = ""
			aOrs = SPLIT( sQuery, "," )
			IF LBOUND(aOrs) < UBOUND(aOrs) THEN
				FOR EACH sOrItem IN aOrs
					aAnds = SPLIT( sOrItem, ":" )
					IF LBOUND(aAnds) < UBOUND(aAnds) THEN
						sBuild = ""
						FOR EACH sAndItem IN aAnds
							sF = categoryQueryBuild( sAndItem )
							IF "" = sBuild THEN
								sBuild = sF
							ELSE
								sBuild = sBuild & " and " & sF
							END IF
						NEXT 'sAndItem
						sF = "(" & sBuild & ")"
					ELSE
						sF = categoryQueryBuild( sOrItem )
					END IF
					IF "" = sBuildMega THEN
						sBuildMega = sF
					ELSE
						sBuildMega = sBuildMega & " or " & sF
					END IF
				NEXT 'sOrItem
				m_sCategoryFilter = "[ " & sBuildMega & " ]"
			ELSE
				aAnds = SPLIT( sQuery, ":" )
				FOR EACH sAndItem IN aAnds
					sF = categoryQueryBuild( sAndItem )
					IF "" = sBuild THEN
						sBuild = sF
					ELSE
						sBuild = sBuild & " and " & sF
					END IF
				NEXT 'sAndItem
				m_sCategoryFilter = "[ " & sBuild & " ]"
			END IF
			'Response.Write "category-filter: " & m_sCategoryFilter
			'Response.Flush
		ELSE
			m_sCategories = ""
			m_sCategoryFilter = ""
		END IF
		
	END PROPERTY
	
	PRIVATE FUNCTION categoryQueryBuild( sCategory )
		IF "~" = LEFT( sCategory, 1 ) THEN
			DIM	sName
			sName = MID(sCategory,2)
			IF "none" = sCategory THEN
				categoryQueryBuild = ""
			ELSE
				categoryQueryBuild = " not( category = '" & sName & "')"
				'categoryQueryBuild = "( $all$ category != """ & sName & """)"
			END IF
		ELSE
			IF "none" = sCategory THEN
				categoryQueryBuild = "not(category)"
			ELSE
				categoryQueryBuild = "( category = '" & sCategory & "')"
				'categoryQueryBuild = "( $any$ category = """ & sCategory & """)"
			END IF
		END IF
	END FUNCTION
	
	
	PRIVATE SUB dateSingle_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		nJulian = jdateFromGregorian( aSplitDate(2), aSplitDate(1), aSplitDate(0) )
		IF nEarly <= nJulian  AND  nJulian <= nLate THEN
			appendList nCount, aDates, nJulian
		END IF
	END SUB
	
	
	PRIVATE SUB dateWeekly_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		DIM	aDayList
		aDayList = SPLIT( aSplitDate(1), ",", -1, vbTextCompare )
		DIM	nWeekday
		nWeekday = weekdayFromJDate( nEarly )
		DIM	nwd
		FOR EACH nwd IN aDayList
			nJulian = nEarly - (CLNG(nWeekday) - 1)
			nJulian = nJulian + (CLNG(nwd) - 1)
			DO WHILE nJulian <= nLate
				IF nEarly <= nJulian THEN
					SELECT CASE CLNG(aSplitDate(0))
					CASE 0
						appendList nCount, aDates, nJulian
					CASE 1
						IF ( (nJulian-6)\7 + 1) MOD 2 = 1 THEN
							appendList nCount, aDates, nJulian
						END IF
					CASE 2
						IF ( (nJulian-6)\7 + 1 ) MOD 2 = 0 THEN
							appendList nCount, aDates, nJulian
						END IF
					END SELECT
				END IF
				nJulian = nJulian + 7
			LOOP
		NEXT
	END SUB
	
	PRIVATE SUB dateMonthlyDayN_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		DIM	dd, mm, cm, cc, yy
		DIM	d, m, y

		gregorianFromJDate d, mm, yy, nEarly
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		cc = CINT(aSplitDate(0))
		cm = CINT(aSplitDate(1))
		dd = CINT(aSplitDate(2))
		IF 1 < cc THEN
			m = mm MOD cc
			m = mm - m + cm
			mm = m
		END IF
		DO
			DO WHILE mm < 13
				nJulian = jdateFromGregorian( dd, mm, yy )
				'Response.Write "julian = " & nJulian & "<br>" & vbCRLF
				IF nLate < nJulian THEN EXIT SUB
				IF 2 = mm  AND  28 < dd THEN
					gregorianFromJDate d, m, y, nJulian
					IF m <> mm  OR  d <> dd THEN nJulian = 0
				END IF
				IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
				mm = mm + cc
			LOOP
			mm = mm MOD 12
			yy = yy + 1
		LOOP
	END SUB
	
	PRIVATE SUB dateMonthlyWDay_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		DIM	yy, cc, cm, mm, wn, wd
		DIM	d, m, y

		gregorianFromJDate d, mm, yy, nEarly
		'Response.Write "date= " & oDate.text & ", year-begin= " & m_nYearBegin & ", yy= " & yy & ", nEarly= " & nEarly & ", nLate= " & nLate & "<br>" & vbCRLF
		'IF 0 = m_nYearBegin THEN
		'	yy = m_nYearNow
		'ELSE
		'	yy = m_nYearBegin
		'END IF
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		cc = CINT(aSplitDate(0))
		cm = CINT(aSplitDate(1))
		wn = CINT(aSplitDate(2))
		wd = CINT(aSplitDate(3))
		IF 1 < cc THEN
			m = mm MOD cc
			m = mm - m + cm
			mm = m
		END IF
		DO
			DO WHILE mm < 13
				nJulian = jdateFromWeeklyGregorian( wn, wd, mm, yy )
				'Response.Write "wn=" & wn & ", wd=" & wd & ", mm=" & mm & ", julian = " & nJulian & "<br>" & vbCRLF
				IF nLate < nJulian THEN EXIT SUB
				IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
				mm = mm + cc
			LOOP
			mm = mm MOD 12
			yy = yy + 1
		LOOP
	END SUB

	PRIVATE SUB dateMonthly_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("monthly")
		CASE "dayn"
			dateMonthlyDayN_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "wday"
			dateMonthlyWDay_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB

	PRIVATE SUB dateYearlyDayN_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		DIM	dd, mm, yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		'IF 0 < UBOUND(aSplitDate) THEN
		mm = CINT(aSplitDate(0))
		dd = CINT(aSplitDate(1))
		DO
			nJulian = jdateFromGregorian( dd, mm, yy )
			'Response.Write "julian = " & nJulian & "<br>" & vbCRLF
			IF nLate < nJulian THEN EXIT DO
			IF 2 = mm  AND  28 < dd THEN
				gregorianFromJDate d, m, y, nJulian
				IF m <> mm  OR  d <> dd THEN nJulian = 0
			END IF
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
		'END IF
	END SUB
	
	PRIVATE SUB dateYearlyWDay_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		DIM	yy, mm, wn, wd
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		'Response.Write "date= " & oDate.text & ", year-begin= " & m_nYearBegin & ", yy= " & yy & ", nEarly= " & nEarly & ", nLate= " & nLate & "<br>" & vbCRLF
		'IF 0 = m_nYearBegin THEN
		'	yy = m_nYearNow
		'ELSE
		'	yy = m_nYearBegin
		'END IF
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		mm = CINT(aSplitDate(0))
		wn = CINT(aSplitDate(1))
		wd = CINT(aSplitDate(2))
		DO
			nJulian = jdateFromWeeklyGregorian( wn, wd, mm, yy )
			'Response.Write "wn=" & wn & ", wd=" & wd & ", mm=" & mm & ", julian = " & nJulian & "<br>" & vbCRLF
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB
		
	PRIVATE SUB dateYearly_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("yearly")
		CASE "dayn"
			dateYearlyDayN_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "wday"
			dateYearlyWDay_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB
	
	PRIVATE SUB dateHebrewDayN_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	aSplitDate
		DIM	nJulian
		DIM	dd, mm, yy
		DIM	d, m, y

		hebrewFromJDate d, m, yy, nEarly
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		'IF 0 < UBOUND(aSplitDate) THEN
		mm = CINT(aSplitDate(0))
		dd = CINT(aSplitDate(1))
		DO
			nJulian = jdateFromHebrew( dd, mm, yy )
			'Response.Write "julian = " & nJulian & "<br>" & vbCRLF
			IF nLate < nJulian THEN EXIT DO
			'IF 2 = mm  AND  28 < dd THEN
			'	gregorianFromJDate d, m, y, nJulian
			'	IF m <> mm  OR  d <> dd THEN nJulian = 0
			'END IF
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
		'END IF
	END SUB

	PRIVATE SUB dateHebrew_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("hebrew")
		CASE "dayn"
			dateHebrewDayN_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "wday"
			dateHebrewWDay_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB
	
	
	PRIVATE FUNCTION dateRoshHashanah_calc( yy )
		dateRoshHashanah_calc = roshhashanah_conway( yy )
	END FUNCTION

	PRIVATE SUB dateRoshHashanah_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = dateRoshHashanah_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB


	PRIVATE FUNCTION datePassover_calc( yy )
		datePassover_calc = passover_conway( yy )
	END FUNCTION

	PRIVATE SUB datePassover_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = datePassover_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB


	PRIVATE FUNCTION dateOrthodoxEaster_calc( yy )
		DIM	d, m
		easter_mallen d, m, yy
		dateOrthodoxEaster_calc = jdateFromGregorian( d, m, yy )
	END FUNCTION

	PRIVATE SUB dateOrthodoxEaster_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = dateOrthodoxEaster_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB


	PRIVATE FUNCTION dateEaster_calc( yy )
		DIM	wd
		DIM	nJulian
		
		nJulian = paschal_moon( yy )
		wd = weekdayFromJDate( nJulian )
		wd = 1 - wd
		IF wd <= 0 THEN wd = wd + 7
		dateEaster_calc = nJulian + wd
	END FUNCTION
	
	PRIVATE SUB dateEaster_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = dateEaster_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB

	PRIVATE SUB datePaschal_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = paschal_moon( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB
	
	PRIVATE SUB dateMoonPhase_calcDateList( nCount, aDates, nPhase, oDate, ByVal nEarly, ByVal nLate )

		IF 0 <= nPhase  AND nPhase <= 3 THEN
			DIM	MOON_CYCLE
			MOON_CYCLE = (29.0+(12.0+(44.0+(2.8/60.0))/60.0)/24.0)
			DIM	n
			DIM nJulian
			DIM	jd
			DIM	frac
			n = FIX((CDBL(nEarly - 2415020) + 1.265465278 _
						+ MOON_CYCLE * (nPhase/4)) / MOON_CYCLE)
			DO
				nJulian = moon_phaseFull( n, nPhase, m_nTimeZoneOffset, jd, frac )
				IF nLate < nJulian THEN EXIT DO
				IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
				n = n + 1
			LOOP
			
		END IF
	END SUB
	
	
	PRIVATE SUB dateMoonFull_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		dateMoonPhase_calcDateList nCount, aDates, 2, oDate, nEarly, nLate
	END SUB


	PRIVATE SUB dateMoonNew_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		dateMoonPhase_calcDateList nCount, aDates, 0, oDate, nEarly, nLate
	END SUB


	PRIVATE SUB dateKeyword_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.text
		CASE "easter"
			dateEaster_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "orthodox easter"
			dateOrthodoxEaster_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "rosh hashanah"
			dateRoshHashanah_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "passover"
			datePassover_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "paschal"
			datePaschal_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "moon full"
			dateMoonFull_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "moon new"
			dateMoonNew_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB
	

	PRIVATE SUB dateSeason_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	dBaseTime
		DIM	dTime
		DIM	nJulian
		SELECT CASE oDate.text
		CASE "0"
			dBaseTime = jtimeFromGregorian( 20, 3, 2000, 7, 35, 0 )
		CASE "1"
			dBaseTime = jtimeFromGregorian( 21, 6, 2000, 1, 48, 0 )
		CASE "2"
			dBaseTime = jtimeFromGregorian( 22, 9, 2000, 17, 27, 0 )
		CASE "3"
			dBaseTime = jtimeFromGregorian( 21, 12, 2000, 13, 37, 0 )
		CASE ELSE
			dBaseTime = 0
		END SELECT

		IF 0 < dBaseTime THEN
			nJulian = nEarly - FIX(dBaseTime)
			nJulian = FIX(nJulian / kTropicalYear)
			dTime = dBaseTime + (kTropicalYear * nJulian)
			IF nEarly < FIX(dTime) THEN dTime = dTime - kTropicalYear
			DO
				nJulian = FIX(dTime)
				IF nLate < nJulian THEN EXIT DO
				IF nEarly <= nJulian THEN appendList nCount, aDates, dTime
				dTime = dTime + kTropicalYear
			LOOP
		END IF
	END SUB
	
	PRIVATE SUB date_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("type")
		CASE "single"
			dateSingle_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "weekly"
			dateWeekly_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "monthly"
			dateMonthly_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "yearly"
			dateYearly_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "hebrew"
			dateHebrew_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "keyword"
			dateKeyword_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "season"
			dateSeason_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "julian"
		CASE ELSE
		END SELECT
	END SUB
	
	PRIVATE FUNCTION offset_calcEarly( nEarly, aOffset() )
		DIM	i
		DIM	nOffset
		
		nOffset = 0
		FOR i = 0 TO UBOUND(aOffset) STEP 2
			SELECT CASE aOffset(i)
			CASE "+", "++", "-", "--"
				nOffset = nOffset + 7
			CASE "#"
				nOffset = nOffset + ABS(CINT(aOffset(i+1)))
			CASE ELSE
			END SELECT
		NEXT 'i
		offset_calcEarly = nEarly - nOffset
	END FUNCTION

	PRIVATE FUNCTION offset_calcLate( nLate, aOffset() )
		DIM	i
		DIM	nOffset
		
		nOffset = 0
		FOR i = 0 TO UBOUND(aOffset) STEP 2
			SELECT CASE aOffset(i)
			CASE "+", "++", "-", "--"
				nOffset = nOffset + 7
			CASE "#"
				nOffset = nOffset - CINT(aOffset(i+1))
			CASE ELSE
			END SELECT
		NEXT 'i
		offset_calcLate = nLate + nOffset
	END FUNCTION
	
	PRIVATE FUNCTION offset_isHoliday( nDate )
		' NEEDS_WORK
		offset_isHoliday = FALSE
	END FUNCTION
	
	PRIVATE FUNCTION offset_evalCondition( nDate, sKwd )
		DIM	wd
		
		offset_evalCondition = FALSE
		wd = weekdayFromJDate( nDate )
		SELECT CASE sKwd
		CASE "weekday"
			IF 1 < wd  AND  wd < 7 THEN offset_evalCondition = TRUE
		CASE "weekend"
			IF wd < 2  OR  6 < wd THEN offset_evalCondition = TRUE
		CASE "workday"
			IF 1 < wd  AND  wd < 7  AND NOT offset_isHoliday( nDate ) THEN
				offset_evalCondition = TRUE
			END IF
		CASE "offday"
			IF wd < 2  OR  6 < wd  OR offset_isHoliday( nDate ) THEN
				offset_evalCondition = TRUE
			END IF
		CASE "holiday"
			IF offset_isHoliday( nDate ) THEN offset_evalCondition = TRUE
		CASE ELSE ' weekday number
			IF IsNumeric(sKwd) THEN
				IF wd = CINT(sKwd) THEN
					offset_evalCondition = TRUE
				END IF
			ELSEIF "m" = LCASE(LEFT(sKwd,1)) THEN
				DIM	dd, mm, yy
				gregorianFromJDate dd, mm, yy, nDate
				wd = MID(sKwd,2)
				IF CINT(wd) = CINT(mm) THEN
					offset_evalCondition = TRUE
				END IF
			END IF
		END SELECT
	END FUNCTION
	
	PRIVATE FUNCTION offset_evalInc( nDate, sKwd, sValue )
		DIM	wd
		DIM	n
		
		wd = weekdayFromJDate( nDate )
		SELECT CASE sValue
		CASE "weekday", "workday"
			IF wd < 2  OR  6 < wd THEN
				SELECT CASE sKwd
				CASE "+", "++"
					IF wd < 2 THEN
						wd = 1
					ELSEIF 6 < wd THEN
						wd = 2
					ELSE
						wd = 0
					END IF
				CASE "-", "--"
					IF wd < 2 THEN
						wd = -2
					ELSEIF 6 < wd THEN
						wd = -1
					ELSE
						wd = 0
					END IF
				CASE "~"
					IF wd < 2 THEN
						wd = 1
					ELSEIF 6 < wd THEN
						wd = -1
					ELSE
						wd = 0
					END IF
				CASE ELSE
					wd = 0
				END SELECT
			ELSE
				wd = 0
			END IF
		CASE "weekend", "offday"
			IF 1 < wd  AND  wd < 7 THEN
				SELECT CASE sKwd
				CASE "+", "++"
					wd = 7 - wd
				CASE "-", "--"
					wd = 1 - wd
				CASE "~"
					IF 3 < wd THEN
						wd = 7 - wd
					ELSE
						wd = 1 - wd
					END IF
				CASE ELSE
					wd = 0
				END SELECT
			ELSE
				wd = 0
			END IF
		CASE "holiday"
			wd = 0
		CASE "cancel"
			wd = -32000
		CASE ELSE
			IF IsNumeric(sValue) THEN
				n = CINT(sValue)
				wd = n - weekdayFromJDate( nDate )
				SELECT CASE sKwd
				CASE "+"
					IF wd < 0 THEN wd = wd + 7
				CASE "++"
					IF wd <= 0 THEN wd = wd + 7
				CASE "-"
					IF 0 < wd THEN wd = wd - 7
				CASE "--"
					IF 0 <= wd THEN wd = wd - 7
				CASE "~"
					IF wd <= -3 THEN
						wd = wd + 7
					ELSEIF 3 < wd THEN
						wd = wd - 7
					END IF
				CASE ELSE
					wd = 0
				END SELECT
			ELSE
				wd = 0
			END IF
		END SELECT
		offset_evalInc = nDate + wd
	END FUNCTION
	
	PRIVATE FUNCTION offset_eval( nDate, aOffset )
		DIM	i
		
		i = 0
		DO WHILE i < UBOUND(aOffset)
			SELECT CASE aOffset(i)
			CASE "?"
				IF NOT offset_evalCondition( nDate, aOffset(i+1) ) THEN i = i + 2
			CASE "+", "++", "-", "--", "~"
				nDate = offset_evalInc( nDate, aOffset(i), aOffset(i+1) )
			CASE "#"
				nDate = nDate + CINT(aOffset(i+1))
			END SELECT
			i = i + 2
		LOOP
		offset_eval = nDate
	END FUNCTION
	
	PRIVATE SUB date_offset_calcDateList( nCount, aDates, oDate, oOffset, ByVal nEarly, ByVal nLate )
		DIM	aSplitTemp
		DIM	sOffset
		DIM aList()
		DIM n, i
		DIM	nOfsEarly, nOfsLate
		DIM	nJulian
		
		sOffset = oOffset.value
		aSplitTemp = SPLIT( sOffset, " ", -1, vbTextCompare )
		nOfsEarly = offset_calcEarly( nEarly, aSplitTemp )
		nOfsLate = offset_calcLate( nLate, aSplitTemp )
		
		REDIM aList(10)
		n = 0
		date_calcDateList n, aList, oDate, nOfsEarly, nOfsLate
		IF 0 < n THEN
			FOR i = 0 TO n
				nJulian = offset_eval( aList(i), aSplitTemp )
				IF nEarly <= nJulian  AND  nJulian <= nLate THEN appendList nCount, aDates, nJulian
			NEXT 'i
		END IF
	END SUB
	
	PRIVATE SUB date_duration( nCount, aDates, ByVal nDur, ByVal nEarly, ByVal nLate )
		DIM	a()
		DIM n, i, j, k
		DIM d
		
		d = nDur
		IF d < 1 THEN d = 1
		REDIM a(nCount)
		FOR i = 0 TO nCount-1
			a(i) = aDates(i)
		NEXT 'i
		n = nCount-1
		nCount = 0
		FOR i = 0 TO n
			j = a(i)
			FOR k = 1 TO d
				IF nEarly <= j  AND  j <= nLate THEN appendList nCount, aDates, j
				j = j + 1
			NEXT 'k
		NEXT 'i
	END SUB
	
	PRIVATE FUNCTION subjectDurationViewBegin( jDate, nDur )
		DIM	s
		DIM	jTarget
		jTarget = jDate - nDur + 1
		IF nDur < 7 THEN	' we will say 'Starts weekday-name'
			s = weekdayNameFromJDate(jTarget, 0)
		ELSE
			DIM	d1, d2
			DIM	m1, m2
			DIM	y1, y2
			gregorianFromJDate d1, m1, y1, jDate
			gregorianFromJDate d2, m2, y2, jTarget
			IF y1 = y2 THEN
				IF m1 = m2 THEN
					s = "on " & d2
					IF 10 < d2  AND  d2 < 20 THEN
						s = s & "th"
					ELSE
						DIM x
						x = RIGHT(CSTR(d2),1)
						SELECT CASE x
						CASE "1"
							s = s & "st"
						CASE "2"
							s = s & "nd"
						CASE "3"
							s = s & "rd"
						CASE ELSE
							s = s & "th"
						END SELECT
					END IF
				ELSE
					s = m2 & "/" & d2
				END IF
			ELSE
				s = m2 & "/" & d2 & "/" & y2
			END IF
		END IF
		subjectDurationViewBegin = " (Starts " & s & ")"
	END FUNCTION
	
	PRIVATE FUNCTION subjectDurationViewEnd( jDate, nDur )
		DIM	s
		DIM	jTarget
		jTarget = jDate + nDur - 1
		IF nDur < 7 THEN	' we will say 'Thru weekday-name'
			s = weekdayNameFromJDate(jTarget, 0)
		ELSE
			DIM	d1, d2
			DIM	m1, m2
			DIM	y1, y2
			gregorianFromJDate d1, m1, y1, jDate
			gregorianFromJDate d2, m2, y2, jTarget
			IF y1 = y2 THEN
				IF m1 = m2 THEN
					s = CSTR(d2)
					IF 10 < d2  AND  d2 < 20 THEN
						s = s & "th"
					ELSE
						DIM x
						x = RIGHT(CSTR(d2),1)
						SELECT CASE x
						CASE "1"
							s = s & "st"
						CASE "2"
							s = s & "nd"
						CASE "3"
							s = s & "rd"
						CASE ELSE
							s = s & "th"
						END SELECT
					END IF
				ELSE
					s = m2 & "/" & d2
				END IF
			ELSE
				s = m2 & "/" & d2 & "/" & y2
			END IF
		END IF
		subjectDurationViewEnd = " (Thru " & s & ")"
	END FUNCTION
	
	PRIVATE FUNCTION subjectDurationViewRange( jDate, nDur )
		DIM	jTarget
		jTarget = jDate + nDur - 1
		DIM	d1, d2
		DIM	m1, m2
		DIM	y1, y2
		DIM	s
		gregorianFromJDate d1, m1, y1, jDate
		gregorianFromJDate d2, m2, y2, jTarget
		IF y2 = y1 THEN
			IF m2 = m1 THEN
				s = d1 & "-" & d2
			ELSE
				s = m1 & "/" & d1 & " - " & m2 & "/" & d2
			END IF
		ELSE
			s = m1 & "/" & d1 & "/" & y1 & " - " & m2 & "/" & d2 & "/" & y2
		END IF
		subjectDurationViewRange = " (" & s & ")"
	END FUNCTION
	
	PRIVATE SUB date_durationView( nCount, aDates, ByVal nDur, ByVal nEarly, ByVal nLate, aAppendSubject )
		DIM	a()
		DIM n, i, j, k
		DIM	jt
		DIM d
		
		d = nDur
		IF d < 1 THEN d = 1
		REDIM a(nCount)
		FOR i = 0 TO nCount-1
			a(i) = aDates(i)
		NEXT 'i
		n = nCount-1
		nCount = 0
		FOR i = 0 TO n
			j = a(i)
			jt = j + d - 1
			IF nEarly <= j  AND  j <= nLate THEN
				IF UBOUND(aDates) < nCount+1 THEN REDIM PRESERVE aDates(nCount+10)
				aDates(nCount) = j
				IF UBOUND(aAppendSubject) < nCount+1 THEN REDIM PRESERVE aAppendSubject(nCount+10)
				aAppendSubject(nCount) = subjectDurationViewEnd( j, d )
				nCount = nCount + 1
			ELSEIF nEarly <= jt  AND jt <= nLate THEN
				IF UBOUND(aDates) < nCount+1 THEN REDIM PRESERVE aDates(nCount+10)
				aDates(nCount) = jt
				IF UBOUND(aAppendSubject) < nCount+1 THEN REDIM PRESERVE aAppendSubject(nCount+10)
				aAppendSubject(nCount) = subjectDurationViewBegin( jt, d )
				nCount = nCount + 1
			ELSEIF j <= nEarly  AND  nLate <= jt  THEN
				IF UBOUND(aDates) < nCount+1 THEN REDIM PRESERVE aDates(nCount+10)
				aDates(nCount) = nEarly
				IF UBOUND(aAppendSubject) < nCount+1 THEN REDIM PRESERVE aAppendSubject(nCount+10)
				aAppendSubject(nCount) = subjectDurationViewRange( j, d )
				nCount = nCount + 1
			END IF
		NEXT 'i
	END SUB
	
	PRIVATE SUB zeroAppendSubject( nCount, aAppendSubject )
		IF UBOUND(aAppendSubject) < nCount THEN
			REDIM aAppendSubject(nCount)
		END IF
		DIM i
		FOR i = 0 TO nCount-1
			aAppendSubject(i) = ""
		NEXT
	END SUB
	
	
	
	

	SUB getDates( pCalendar )
	
		DIM	oEvents
		DIM	oEvent
		DIM	oCategories
		DIM	oCategory
		DIM	oDates
		DIM	oDate
		DIM	oItem
		DIM	bProcess
		
		DIM	sID
		DIM	sStyle
		DIM	sSubject
		DIM	sContent
		DIM	sLocation
		DIM	sComments
		DIM	sTime
		DIM	sSingle
		DIM	nJulian
		DIM	nDuration
		DIM	sDurationView
		DIM	nPending
		DIM	nEarly
		DIM	nLate
		DIM i
		
		DIM	aDates()
		DIM	aAppendSubjects()
		DIM	nDatesCount
		
		REDIM aDates(20)
		REDIM aAppendSubjects(20)
		
		
		IF m_oXML IS Nothing THEN
			Response.Write "Remind - xml document not loaded<br>"
			EXIT SUB
		END IF
				
		
		'SET oEvents = m_oXML.selectNodes("//event")
		SET oEvents = m_oXML.selectNodes("//event" & m_sCategoryFilter)
		'Response.Write oEvents.length & "<br>"
		FOR EACH oEvent IN oEvents
			SET oItem = oEvent.attributes.getNamedItem("id")
			IF NOT Nothing IS oItem THEN
				sID = oItem.value
			ELSE
				sID = ""
			END IF

			SET oItem = oEvent.selectSingleNode("subject")
			sSubject = oItem.text
			SET oItem = Nothing
			ON Error RESUME Next
			SET oItem = oEvent.selectSingleNode("location")
			IF NOT oItem IS Nothing THEN
				sLocation = oItem.text
			ELSE
				sLocation = ""
			END IF
			SET oItem = oEvent.selectSingleNode("comments")
			IF NOT oItem IS Nothing THEN
				sComments = oItem.text
			ELSE
				sComments = ""
			END IF
			ON Error GOTO 0
			sStyle = ""
			ON Error RESUME Next
			SET oItem = oEvent.selectSingleNode("style")
			IF NOT oItem IS Nothing THEN sStyle = oItem.text
			ON Error GOTO 0
			sContent = ""
			ON Error RESUME Next
			SET oItem = oEvent.selectSingleNode("content")
			IF NOT oItem IS Nothing THEN sContent = oItem.text
			ON Error GOTO 0
			SET oDates = oEvent.getElementsByTagName("date")
			nDatesCount = 0
			FOR EACH oDate IN oDates
				sTime = ""
				SET oItem = oDate.attributes.getNamedItem("time")
				IF NOT Nothing IS oItem THEN
					sTime = oItem.value
				END IF
				sDurationView = ""
				nDuration = 0
				nEarly = m_nDateBegin
				nLate = m_nDateEnd
				IF 0 < m_nPending THEN
					nPending = 0
					SET oItem = oDate.attributes.getNamedItem("pending")
					IF NOT( Nothing IS oItem ) THEN
						nPending = CLNG(oItem.value)
						nLate = nEarly + nPending
					END IF
				END IF
				SET oItem = oDate.attributes.getNamedItem("duration")
				IF NOT( Nothing IS oItem ) THEN
                    IF ISNUMERIC(oItem.value) THEN
					    nDuration = CLNG(oItem.value)
					    SET oItem = oDate.attributes.getNamedItem("duration-view")
					    IF NOT oItem IS Nothing THEN
						    sDurationView = oItem.value
					    END IF
					    nEarly = nEarly - nDuration
                    END IF
				END IF
				SET oItem = oDate.attributes.getNamedItem("offset")
				IF Nothing IS oItem THEN
					date_calcDateList nDatesCount, aDates, oDate, nEarly, nLate
				ELSE
					date_offset_calcDateList nDatesCount, aDates, oDate, oItem, nEarly, nLate
				END IF
				IF 1 < nDuration THEN
					nEarly = nEarly + nDuration
					IF "" <> sDurationView THEN
						IF sDurationView = m_sCalendarView THEN
							date_duration nDatesCount, aDates, nDuration, nEarly, nLate
							zeroAppendSubject nDatesCount, aAppendSubjects
						ELSE
							date_durationView nDatesCount, aDates, nDuration, nEarly, nLate, aAppendSubjects
						END IF
					ELSE
						date_duration nDatesCount, aDates, nDuration, nEarly, nLate
						zeroAppendSubject nDatesCount, aAppendSubjects
					END IF
				ELSE
					zeroAppendSubject nDatesCount, aAppendSubjects
				END IF
				SET oCategories = oEvent.selectNodes("category")

				IF 0 < nDatesCount THEN
					FOR i = 0 TO nDatesCount-1
						nJulian = FIX(aDates(i))
						IF nEarly <= nJulian  AND  nJulian <= nLate THEN
							IF 0 < LEN(sID) THEN pCalendar.id = sID
							pCalendar.juliandate = aDates(i)
							pCalendar.subject = sSubject & aAppendSubjects(i)
							IF 0 < LEN(sTime) THEN pCalendar.time = sTime
							IF 0 < LEN(sContent) THEN pCalendar.content = sContent
							IF 0 < LEN(sLocation) THEN pCalendar.location = sLocation
							IF 0 < LEN(sComments) THEN pCalendar.comments = sComments
							IF 0 < LEN(sStyle) THEN pCalendar.style = sStyle
							FOR EACH oCategory IN oCategories
								pCalendar.category = oCategory.text
							NEXT 'oCategory
							pCalendar.outputMessage
						END IF
					NEXT 'i
					nDatesCount = 0
				END IF

			NEXT 'oDate
		NEXT 'oEvent
		
	END SUB
	
END CLASS 'CCalendarFile



'===========================================================



%>