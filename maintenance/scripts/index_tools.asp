<%



SUB locationSort( A(), ByVal Lb, ByVal Ub )
    DIM n
    DIM h
    DIM i
    DIM j
    DIM t

    ' sort array[lb..ub]

    ' compute largest increment
    n = Ub - Lb + 1
    h = 1
    IF (n < 14) THEN
        h = 1
    ELSE
        DO WHILE h < n
            h = 3 * h + 1
        LOOP
        h = h \ 3
        h = h \ 3
    END IF

    DO WHILE h > 0
        ' sort by insertion in increments of h
        FOR i = Lb + h TO Ub
            t = A(i)
            FOR j = i - h TO Lb STEP -h
                IF A(j) <= t THEN EXIT FOR
                A(j + h) = A(j)
            NEXT 'j
            A(j + h) = t
        NEXT 'i
        h = h \ 3
    LOOP
END SUB



SUB locationSortDescend( A(), ByVal Lb, ByVal Ub )
    DIM n
    DIM h
    DIM i
    DIM j
    DIM t
	DIM	t2

    ' sort array[lb..ub]

    ' compute largest increment
    n = Ub - Lb + 1
    h = 1
    IF (n < 14) THEN
        h = 1
    ELSE
        DO WHILE h < n
            h = 3 * h + 1
        LOOP
        h = h \ 3
        h = h \ 3
    END IF

    DO WHILE h > 0
        ' sort by insertion in increments of h
        FOR i = Lb + h TO Ub
            t = A(i)
            FOR j = i - h TO Lb STEP -h
                IF t <= A(j) THEN EXIT FOR
                A(j + h) = A(j)
            NEXT 'j
            A(j + h) = t
        NEXT 'i
        h = h \ 3
    LOOP
END SUB



FUNCTION HTMLDecode( sString )
	DIM	sTemp
	
	IF 0 < LEN(sString) THEN
		IF INSTR( 1, sString, "&", vbTextCompare ) THEN
			sTemp = REPLACE( sString, "&amp;", "&", 1, -1, vbTextCompare )
			sTemp = REPLACE( sTemp, "&lt;", "<", 1, -1, vbTextCompare )
			sTemp = REPLACE( sTemp, "&gt;", ">", 1, -1, vbTextCompare )
			sTemp = REPLACE( sTemp, "&quot;", """", 1, -1, vbTextCompare )
			HTMLDecode = REPLACE( sTemp, "&nbsp;", " ", 1, -1, vbTextCompare )
		ELSE
			HTMLDecode = sString
		END IF
	ELSE
		HTMLDecode = ""
	END IF
END FUNCTION


DIM	g_sNavigateLabel
DIM	g_oNavigateReg
g_sNavigateLabel = ",tab,"


SUB identifyNavigateLabel()
	DIM	n
	ON ERROR Resume Next
	n = VARTYPE(g_sNavigateTabLabel)
	IF 0 = Err THEN
		IF 8 = n THEN	'type=string
			g_sNavigateLabel = "," & g_sNavigateTabLabel & ","
		END IF
	END IF
	ON ERROR Goto 0
	SET g_oNavigateReg = NEW RegExp
	g_oNavigateReg.Pattern = g_sNavigateLabel
	g_oNavigateReg.IgnoreCase = TRUE
	g_oNavigateReg.Global = TRUE
END SUB
identifyNavigateLabel


SUB getWebFileInfo( sTitle, sDescription, bIsIndex, _
					sTargetWindow, sSortName, _
					sDateRange, sEventDate, sKeywords, _
					sFile )
	DIM objInFile			'object variables for file access
	DIM strIn				'string variables for reading
	DIM	strTemp
	DIM	i,j
	DIM bProcessString		'flag determining whether or not to work with each line
	DIM	bInTitle
	DIM	bRobots
	DIM	nMetaNav
	
	bRobots = TRUE
	nMetaNav = 0	' neutral=0, no=-1, yes=1

	bIsIndex = FALSE
	bInTitle = FALSE
	
	sTitle = ""
	sDescription = ""
	sTargetWindow = ""
	sSortName = ""
	sDateRange = ""
	sEventDate = ""
	sKeywords = ""

	bProcessString = 0

	IF sFile <> "" THEN
	
		DIM	sDescText
		DIM	sComments
		DIM	sStatus
		DIM	sVersion
		DIM	sAuthor
		DIM	sSubject
		DIM	sNavTitle
		DIM	sKwy
		
		sDescText = ""
		sAuthor = ""
		sSubject = ""
		sComments = ""
		sStatus = ""
		sVersion = ""
		sNavTitle = ""
		sKwy = ""

		' we are going to trust that the file exists since it was enumerated
		'IF g_oFSO.FileExists( sFile ) THEN

			'Response.Write sFile & "<br>" & vbCRLF
			SET objInFile = g_oFSO.OpenTextFile( sFile )
			
			DIM	sFullContent
			sFullContent = objInFile.ReadAll()
			
			DIM	oReg
			DIM	oMatchList
			DIM	oMatch
			
			SET oReg = NEW RegExp
			oReg.Pattern = "(<head>[\p\P\s\S\w\W]*</head>)"
			oReg.IgnoreCase = TRUE
			oReg.Global = TRUE
			SET oMatchList = oReg.Execute( sFullContent )
			IF 0 < oMatchList.Count THEN
				SET oMatch = oMatchList(0)
				
				DIM	sTemp
				DIM	sHead
				sHead = TRIM(oMatch.SubMatches(0))
				
				oReg.Pattern = "<title>(..*)</title>"
				SET oMatchList = oReg.Execute( sHead )
				IF 0 < oMatchList.Count THEN
					sTemp = oMatchList(0).SubMatches(0)
					sTemp = REPLACE(sTemp,vbCRLF,vbLF)
					sTemp = REPLACE(sTemp,vbCR,vbLF)
					sTemp = REPLACE(sTemp,vbLF," ")
					sTemp = REPLACE(sTemp,"  "," ")
					sTitle = TRIM(sTemp)
				END IF
				
				oReg.Pattern = "<meta\s+(name|http-equiv)=""([^""]*)""\s+content=""([^""]*)""\s*[/]?>"
				oReg.IgnoreCase = TRUE
				oReg.Global = TRUE
				SET oMatchList = oReg.Execute( sHead )
				IF 0 < oMatchList.Count THEN
				
					DIM	sName
					DIM	sContent
				
					FOR EACH oMatch IN oMatchList
					
						sName = LCASE(oMatch.SubMatches(1))
						sContent = TRIM(oMatch.SubMatches(2))
						'Response.Write sName & " - " & sContent & "<br>" & vbCRLF
						
						SELECT CASE sName
						CASE "robots"
							sTemp = ","&LCASE(sContent)&","
							IF 0 < INSTR(sTemp, ",none,") THEN
								bRobots = FALSE
								EXIT FOR
							ELSEIF 0 < INSTR(sTemp, ",noindex,") THEN
								bRobots = FALSE
								EXIT FOR
							END IF

						CASE "navigate"
							sTemp = ","&LCASE(sContent)&","
							IF 0 < INSTR(sTemp, ",none,") THEN
								nMetaNav = -1
								bIsIndex = FALSE
								EXIT FOR
							ELSEIF 0 < INSTR(sTemp, ",noindex,") THEN
								nMetaNav = -1
								bIsIndex = FALSE
								EXIT FOR
							ELSEIF g_oNavigateReg.Test(sTemp) THEN
								nMetaNav = 1
								bIsIndex = TRUE
							END IF

						CASE "navtitle"
							sNavTitle = sContent
							
						CASE "sortname"
							sSortName = LCASE(sContent)
							
						CASE "description"
							sDescText = sContent
							
						CASE "window-target", "target-window"
							sTargetWindow = LCASE(sContent)
						
						CASE "daterange", "date-range"
							sDateRange = TRIM(sContent)
						
						CASE "eventdate", "event-date", "dateevent"
							sEventDate = TRIM(sContent)
						
						CASE "keywords"
							sKwy = TRIM(sContent)
							sKeywords = REPLACE(sKwy,", ", ",")
						
						END SELECT
					
					NEXT 'oMatch
					
				END IF
				
				SET oMatch = Nothing
				SET oMatchList = Nothing
				SET oReg = Nothing
			
			END IF
			
			' Close file and free variables
			objInFile.Close
			SET objInFile = Nothing
			
			IF "" <> sNavTitle THEN sTitle = sNavTitle
			sTitle = HTMLDecode(sTitle)
			
			sDescription = HTMLDecode(sDescText)

		'ELSE

			'Response.Write "file not found: " & strFileName

		'END IF

	END IF
END SUB





DIM g_sUseFileNameSuffix
g_sUseFileNameSuffix = ".asp"

FUNCTION useFileName( sFileName )

	DIM	sSuffix
	DIM	i
	
	useFileName = FALSE
	i = InStrRev( sFileName, ".", -1, vbTextCompare )
	IF 0 < i THEN
		sSuffix = MID( sFileName, i )
		IF g_sUseFileNameSuffix = LCASE(sSuffix) THEN
			IF "_" = LEFT( sFileName, 1 ) THEN
				useFileName = FALSE
			ELSE
				useFileName = TRUE
			END IF
		END IF
	END IF

END FUNCTION


FUNCTION genSortnameDeltaFromNow( d )

	genSortnameDeltaFromNow = ""

	IF ISDATE( d ) THEN
		DIM dNowTZ
		dNowTZ = DATEADD("h", g_nServerTimeZoneOffset, NOW )
		DIM n
		n = DATEDIFF( "d", d, dNowTZ )
		IF 7 < n THEN n = n * 2.5
		genSortnameDeltaFromNow = genSortnameFromDate( DATEADD( "d", n, dNowTZ ) )
	END IF

END FUNCTION


FUNCTION genSortnameFromDate( d )

	genSortnameFromDate = ""
	DIM	dd
	DIM	mm
	DIM	yy
	DIM	dDate
	
	IF ISDATE( d ) THEN
		dDate = CDATE( d )

		dd = DAY( dDate )
		mm = MONTH( dDate )
		yy = YEAR( dDate )
		
		genSortnameFromDate = CSTR(yy) & RIGHT("0"&mm,2) & RIGHT("0"&dd,2)

	END IF

END FUNCTION


CONST kFI_SortName = 0
CONST kFI_Name = 1
CONST kFI_Path = 2
CONST kFI_Title = 3
CONST kFI_Description = 4
CONST kFI_TargetWindow = 5
CONST kFI_DateLastModified = 6
CONST kFI_DateRange = 7
CONST kFI_EventDate = 8
CONST kFI_Keywords = 9
CONST kFI_Image = 10

DIM aFileList()
DIM nFileCount
REDIM aFileList(5)
nFileCount = 0

SUB buildFileList( sPath, bDynSortname )

	DIM oFolder
	DIM oFile
	DIM sName
	DIM sTitle
	DIM	sDescription
	DIM sTargetWindow
	DIM	sDateRange
	DIM	sEventDate
	DIM bIsIndex
	DIM	sSortName
	DIM	sKeywords
	
	DIM	i
	DIM	j
	DIM	sDateBegin
	DIM	sDateEnd
	DIM	dNowTZ
	
	dNowTZ = DATEADD("h", g_nServerTimeZoneOffset, NOW )
	dNowTZ = DATEVALUE(dNowTZ)
	
	nFileCount = 0

	SET oFolder = g_oFSO.GetFolder( Server.MapPath( sPath ) )
	
	FOR EACH oFile IN oFolder.Files
		sName = LCase( oFile.Name )
		IF useFileName( sName ) THEN
			getWebFileInfo sTitle, sDescription, bIsIndex, sTargetWindow, sSortName, sDateRange, sEventDate, sKeywords, oFile.Path
			'Response.Write "IsIndex=" & bIsIndex & ": " & sName & "<br>" & vbCRLF
			IF "" <> sDateRange THEN
				i = INSTR(sDateRange,":")
				IF 0 < i THEN
					sDateBegin = LEFT(sDateRange,i-1)
					sDateEnd = MID(sDateRange,i+1)
				ELSE
					sDateBegin = DATEVALUE(DATEADD("h", g_nServerTimeZoneOffset, oFile.DateCreated))
					sDateEnd = sDateRange
				END IF
				IF ISDATE(sDateBegin)  AND  ISDATE(sDateEnd) THEN
					IF dNowTZ < CDATE(sDateBegin) THEN
						bIsIndex = FALSE
					ELSEIF CDATE(sDateEnd) < dNowTZ THEN
						bIsIndex = FALSE
					END IF
				END IF
			ELSE
				sDateBegin = DATEVALUE(DATEADD("h", g_nServerTimeZoneOffset, oFile.DateCreated))
			END IF
			IF bIsIndex THEN
				IF UBOUND(aFileList) <= nFileCount THEN
					REDIM PRESERVE aFileList(nFileCount+20)
				END IF
				IF "" <> sSortName THEN
					sSortName = LCASE(sSortName) & sName
				ELSE
					IF bDynSortname THEN
						IF "" <> sEventDate THEN
							IF ISDATE(sEventDate) THEN
								sSortName = genSortnameFromDate( sEventDate )
								IF dNowTZ <= CDATE(sEventDate) THEN
									i = DATEDIFF( "d", dNowTZ, CDATE(sEventDate) )
									IF i < 7 THEN
										sSortName = STRING( 7-i, "0" ) & sSortName
									ELSE
										j = DATEDIFF( "d", CDATE(sDateBegin), dNowTZ )
										IF j < 8 THEN
											sSortName = genSortnameDeltaFromNow( CDATE(sDateBegin) )
										END IF
									END IF
								END IF
							END IF
						ELSE
							sSortName = genSortnameDeltaFromNow( CDATE(sDateBegin) )
						END IF
					END IF
					sSortName = sSortName & sName
				END IF
				aFileList(nFileCount) = sSortName _
						& vbTAB & sName _
						& vbTAB & oFile.Path _
						& vbTAB & sTitle _
						& vbTAB & sDescription _
						& vbTAB & sTargetWindow _
						& vbTAB & oFile.DateLastModified _
						& vbTAB & sDateRange _
						& vbTAB & sEventDate _
						& vbTAB & sKeywords _
						& vbTAB & ""
				'Response.Write Server.HTMLEncode(aFileList(nFileCount)) & "<br>" & vbCRLF
				nFileCount = nFileCount + 1
			END IF
		END IF
	NEXT 'oFile
	
	IF 0 < nFileCount THEN
		locationSort aFileList, 0, nFileCount-1
		REDIM PRESERVE aFileList(nFileCount-1)
	END IF
	
	'IF bDynSortname THEN
	'	FOR EACH sName IN aFileList
	'		Response.Write Server.HTMLEncode(sName) & "<br>"
	'	NEXT
	'END IF

END SUB

buildFileList "./", FALSE



PUBLIC aFileSplit
DIM g_sPage
DIM g_sPageTitle
DIM g_sFile
DIM g_nIndex
DIM nLen

g_sFile = ""
g_nIndex = -1
g_sPage = Request.ServerVariables("PATH_TRANSLATED")
nLen = INSTRREV(g_sPage,"\",-1,vbTextCompare)
IF 0 < nLen THEN g_sPage = MID(g_sPage,nLen+1)
FOR nLen = 0 TO nFileCount-1
	aFileSplit = SPLIT( aFileList(nLen), vbTAB, -1, vbTextCompare )
	IF LCASE(g_sPage) = LCASE(aFileSplit(kFI_Name)) THEN
		g_sPageTitle = aFileSplit(kFI_Title)
		g_sFile = aFileSplit(kFI_Path)
		g_nIndex = nLen
		EXIT FOR
	END IF
NEXT 'nLen
IF 0 = LEN(g_sFile) THEN
	IF 0 < nFileCount THEN
		aFileSplit = SPLIT( aFileList(0), vbTAB, -1, vbTextCompare )
		g_sPage = aFileSplit(kFI_Name)
		g_sPageTitle = aFileSplit(kFI_Title)
		g_sFile = aFileSplit(kFI_Path)
		g_nIndex = 0
	ELSE
		g_sPage = "?????"
		g_sPageTitle = "??????"
		g_sFile = "?????"
		g_nIndex = 0
	END IF
END IF


CLASS CTabDataFiles

	PRIVATE m_aData
	PRIVATE m_aSplit
	PRIVATE m_i
	PRIVATE m_sURL

	PRIVATE SUB Class_Initialize
		m_sURL = ""
		m_i = 0
	END SUB

	
	PROPERTY GET RecordCount()
		RecordCount = UBOUND(m_aData) + 1
	END PROPERTY
	
	PROPERTY GET EOF()
		IF m_i <= UBOUND(m_aData) THEN
			EOF = FALSE
		ELSE
			EOF = TRUE
		END IF
	END PROPERTY
	
	SUB MoveFirst()
		m_i = 0
		privParse
	END SUB
	
	SUB MoveNext()
		m_i = m_i + 1
		privParse
	END SUB
	
	FUNCTION IsCurrent( x )
		IF m_i = x THEN
			IsCurrent = TRUE
		ELSE
			IsCurrent = FALSE
		END IF
	END FUNCTION
	
	PROPERTY GET HREF()
		HREF = m_aSplit(kFI_Name)
	END PROPERTY
	
	PROPERTY GET TabLabel()
		TabLabel = m_aSplit(kFI_Title)
	END PROPERTY
	
	'----------------
	
	PRIVATE SUB privParse()
		IF NOT EOF() THEN
			m_aSplit = SPLIT( m_aData(m_i), vbTAB, -1, vbTextCompare )
		END IF
	END SUB
	
	PROPERTY LET Data( a )
		m_aData = a
	END PROPERTY
	
	PROPERTY LET URL( s )
		m_sURL = s
	END PROPERTY
	
END CLASS






%>