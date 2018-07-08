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

CONST GREGORIAN_EPOCH = 1721425.5
CONST HEBREW_EPOCH = 347995.5

FUNCTION dateMOD( a, b )
    dateMOD = a - (b * FIX(a / b))
END FUNCTION


FUNCTION gregorianIsLeap( y )

	gregorianIsLeap = FALSE
	IF 0 = (y MOD 4) THEN
		IF NOT( 0 = (y MOD 100)  AND  0 <> (y MOD 400) ) THEN
			gregorianIsLeap = TRUE
		END IF
	END IF

END FUNCTION


FUNCTION hebrewIsLeap( y )
	IF dateMOD((( y * 7 ) + 1 ), 19) < 7 THEN
		hebrewIsLeap = TRUE
	ELSE
		hebrewIsLeap = FALSE
	END IF
END FUNCTION


FUNCTION hebrewMonthsInYear( y )
	IF hebrewIsLeap( y ) THEN
		hebrewMonthsInYear = 13
	ELSE
		hebrewMonthsInYear = 12
	END IF
END FUNCTION


''  Test for delay of start of new year and to avoid
''  Sunday, Wednesday, and Friday as start of the new year.

FUNCTION hebrewDelay1( y )

    DIM months, days, parts

    months = FIX(((235 * y) - 234) / 19)
    parts = 12084 + (13753 * months)
    days = (months * 29) + FIX(parts / 25920)

    IF dateMOD((3 * (days + 1)), 7) < 3 THEN
    	days = days + 1
    END IF
    hebrewDelay1 = days

END FUNCTION


''  Check for delay in start of new year due to length of adjacent years

FUNCTION hebrewDelay2( y )

    DIM last, present, nxt

    last = hebrewDelay1(y - 1)
    present = hebrewDelay1(y)
    nxt = hebrewDelay1(y + 1)
    
    IF 356 = nxt - present THEN
    	hebrewDelay2 = 2
    ELSE
    	IF 382 = present - last THEN
    		hebrewDelay2 = 1
    	ELSE
    		hebrewDelay2 = 0
    	END IF
    END IF

END FUNCTION


''  How many days are in a Hebrew year ?

FUNCTION hebrewDaysInYear( y )
	hebrewDaysInYear = jdateFromHebrew( 1, 7, y+1 ) - jdateFromHebrew( 1, 7, y )
END FUNCTION


''  How many days are in a given month of a given year

FUNCTION hebrewDaysInMonth( m, y )

    '  First of all, dispose of fixed-length 29 day months

    IF m = 2 OR m = 4 OR m = 6 OR m = 10 OR m = 13 THEN
    	hebrewDaysInMonth = 29
    ELSEIF m = 12  AND NOT hebrewIsLeap( y ) THEN
    	'  If it's not a leap year, Adar has 29 days
    	hebrewDaysInMonth = 29
    ELSEIF m = 8  AND  NOT( dateMOD( hebrewDaysInYear( y ), 10 ) = 5 ) THEN
		'  If it's Heshvan, days depend on length of year
    	hebrewDaysInMonth = 29
    ELSEIF m = 9  AND  NOT( dateMOD( hebrewDaysInYear( y ), 10 ) = 3 ) THEN
    	'  Similarly, Kislev varies with the length of year
    	hebrewDaysInMonth = 29
	ELSE
    	'  Nope, it's a 30 day month
    	hebrewDaysInMonth = 30
	END IF

END FUNCTION


FUNCTION jdateFromHebrew( d, m, y )

	DIM jd, mon, months

    months = hebrewMonthsInYear(y)
    jd = HEBREW_EPOCH + hebrewDelay1(y) + hebrewDelay2(y) + d + 1

	IF m < 7 THEN
		FOR mon = 7 TO months
			jd = jd + hebrewDaysInMonth( mon, y )
		NEXT 'mon
		FOR mon = 1 TO m-1
			jd = jd + hebrewDaysInMonth( mon, y )
		NEXT 'mon
	ELSE
		FOR mon = 7 TO m - 1
			jd = jd + hebrewDaysInMonth( mon, y )
		NEXT 'mon
	END IF
	
	jdateFromHebrew = jd + 0.5

END FUNCTION


SUB hebrewFromJDate( d, m, y, jd )
	DIM	ijd
	DIM year, month, day, i, count, first

    ijd = FIX(jd - 0.5) '+ 0.5
    count = FIX(((ijd - HEBREW_EPOCH) * 98496.0) / 35975351.0)
    y = count - 1
    i = count
    DO WHILE jdateFromHebrew(1, 7, i) <= jd
    	y = y + 1
    	i = i + 1
    LOOP
    IF jd < jdateFromHebrew( 1, 1, y ) THEN
    	first = 7
    ELSE
    	first = 1
    END IF
    m = first
    i = first
    DO WHILE jdateFromHebrew( hebrewDaysInMonth( i, y ), i, y ) < jd
    	m = m + 1
    	i = i + 1
    LOOP
    d = (jd - jdateFromHebrew(1, m, y)) + 1

END SUB






SUB gregorianFromJDate( d, m, y, jd )

    DIM wjd, depoch, quadricent, dqc, cent, dcent, quad, dquad, yindex, dyindex, yearday, leapadj

    wjd = FIX(jd - 0.5) '+ 0.5
    wjd = jd
    depoch = wjd - GREGORIAN_EPOCH
    quadricent = FIX(depoch / 146097)
    dqc = dateMOD(depoch, 146097)
    cent = FIX(dqc / 36524)
    dcent = dateMOD(dqc, 36524)
    quad = FIX(dcent / 1461)
    dquad = dateMOD(dcent, 1461)
    yindex = FIX(dquad / 365)
    y = (quadricent * 400) + (cent * 100) + (quad * 4) + yindex
    IF NOT((cent = 4) OR (yindex = 4)) THEN
        y = y + 1
    END IF
    yearday = wjd - jdateFromGregorian(1, 1, y)
    IF wjd < jdateFromGregorian(1, 3, y) THEN
    	leapadj = 0
    ELSE
		IF gregorianIsLeap(y) THEN
			leapadj = leapadj + 1
		ELSE
			leapadj = leapadj + 2
		END IF
    END IF
    m = FIX((((yearday + leapadj) * 12) + 373) / 367)
    d = (wjd - jdateFromGregorian(1, m, y)) + 1

END SUB




FUNCTION jdateFromGregorian( d, m, y )

	DIM	jmon
	DIM	jdate
	DIM	dd, mm, yy

	dd = CLNG(d)
	mm = CLNG(m)
	yy = CLNG(y)

	IF dd < 0 THEN
		dd = dd + 1
		mm = mm + 1
		IF 12 < mm THEN
			mm = mm - 12
			yy = yy + 1
		END IF
	END IF
	
	IF mm <= 2 THEN
		jmon = 0
	ELSE
		IF gregorianIsLeap(yy) THEN
			jmon = -1
		ELSE
			jmon = -2
		END IF
	END IF

    jdate = (GREGORIAN_EPOCH - 1) _
    	+ (365 * (yy - 1)) _
    	+ FIX((yy - 1) / 4) _
    	+ (-FIX((yy - 1) / 100)) _
    	+ FIX((yy - 1) / 400) _
    	+ FIX((((367 * mm) - 362) / 12) _
    	+ jmon _
    	+ dd)
	jdateFromGregorian = jdate + 0.5

END FUNCTION


FUNCTION jtimeFromGregorian( d, m, y, hr, mn, sc )
	DIM	jdate
	DIM	nSec
	
	jdate = jdateFromGregorian( d, m, y )
	nSec = sc + mn*60 + hr*3600
	jtimeFromGregorian = CDBL(jdate) + (CDBL(nSec) / CDBL(3600*24))
	
END FUNCTION


DIM k_dateutils_weekdaynames
k_dateutils_weekdaynames = ARRAY( "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" )


FUNCTION weekdayNameFromNumber( nWD, nAbbrev )
	IF 1 <= nWD  AND  nWD <= 7 THEN
		IF 0 < nAbbrev THEN
			weekdayNameFromNumber = LEFT(k_dateutils_weekdaynames(nWD-1), nAbbrev)
		ELSE
			weekdayNameFromNumber = k_dateutils_weekdaynames(nWD-1)
		END IF
	ELSE
		weekdayNameFromNumber = ""
	END IF
END FUNCTION


' 1=sunday
FUNCTION weekdayFromJDate( jd )
    'weekdayFromJDate = dateMOD(FIX((jd + 1.5)), 7) + 1

	weekdayFromJDate = (jd - 6) MOD 7 + 1

END FUNCTION


FUNCTION weekdayNameFromJDate( jd, nAbbrev )
	weekdayNameFromJDate = weekdayNameFromNumber( weekdayFromJDate(jd), nAbbrev )
END FUNCTION


FUNCTION jdateFromWeeklyGregorian( wn, wday, m, y )

	DIM	jdate
	DIM	n
	DIM	weeknum, wd, mm, yy
	
	weeknum = wn
	wd = wday
	mm = m
	yy = y

	IF weeknum < 0 THEN
		mm = mm + 1
		IF 12 < mm THEN
			mm = 1
			yy = yy + 1
		END IF
		weeknum = weeknum + 1
	END IF

	'
	'	first let's come up with the julian date of the first of the
	'	month for the indicated month.
	'
	jdate = jdateFromGregorian( 1, mm, yy )
	n = weekdayFromJDate( jdate )
	n = wd - n
	IF n < 0 THEN n = n + 7
	jdate = jdate + n + (( weeknum - 1 ) * 7)
	
	'
	'	now that we have a julian date let's make sure that it really
	'	falls in the indicated month.
	'
	IF 0 < weeknum THEN
		mm = mm + 1
		IF 12 < mm THEN
			mm = 1
			yy = yy + 1
		END IF
		IF jdateFromGregorian( 1, mm, yy ) <= jdate THEN jdate = 0	' problem
	ELSE
		mm = mm - 1
		IF mm < 1 THEN
			mm = 12
			yy = yy - 1
		END IF
		IF jdate < jdateFromGregorian( 1, mm, yy ) THEN jdate = 0	' problem
	END IF

	jdateFromWeeklyGregorian = jdate

END FUNCTION


FUNCTION diffWeekdaysJDates( jdFrom , jdTo, bInclEnd )
	DIM	nDiff
	DIM	nOffset
	DIM	nFromWD
	DIM	nToWD
	DIM	nNextSat
	DIM	nNextSun
	
	nDiff = jdTo - jdFrom
	IF bInclEnd THEN
		nDiff = nDiff + 1
		nOffset = 0
	ELSE
		nOffset = 1
	END IF
	nFromWD = weekdayFromJDate( jdFrom )
	nToWD = weekdayFromJDate( jdTo )
	nNextSat = jdFrom + (7 - nFromWD)
	nNextSun = jdFrom - nFromWD + 1
	IF 1 < nFromWD THEN nNextSun = nNextSun + 7
	diffWeekdaysJDates = nDiff
	IF nNextSat <= jdTo-nOffset THEN
		diffWeekdaysJDates = diffWeekdaysJDates - ((jdTo - nNextSat - nOffset) \ 7 + 1)
	END IF
	IF nNextSun <= jdTo-nOffset THEN
		diffWeekdaysJDates = diffWeekdaysJDates - ((jdTo - nNextSun - nOffset) \ 7 + 1)
	END IF
END FUNCTION



DIM aPaschalTable(2,19)
aPaschalTable(0,0) = 4
aPaschalTable(1,0) = 14
aPaschalTable(0,1) = 4
aPaschalTable(1,1) = 3
aPaschalTable(0,2) = 3
aPaschalTable(1,2) = 23
aPaschalTable(0,3) = 4
aPaschalTable(1,3) = 11
aPaschalTable(0,4) = 3
aPaschalTable(1,4) = 31
aPaschalTable(0,5) = 4
aPaschalTable(1,5) = 18
aPaschalTable(0,6) = 4
aPaschalTable(1,6) = 8
aPaschalTable(0,7) = 3
aPaschalTable(1,7) = 28
aPaschalTable(0,8) = 4
aPaschalTable(1,8) = 16
aPaschalTable(0,9) = 4
aPaschalTable(1,9) = 5
aPaschalTable(0,10) = 3
aPaschalTable(1,10) = 25
aPaschalTable(0,11) = 4
aPaschalTable(1,11) = 13
aPaschalTable(0,12) = 4
aPaschalTable(1,12) = 2
aPaschalTable(0,13) = 3
aPaschalTable(1,13) = 22
aPaschalTable(0,14) = 4
aPaschalTable(1,14) = 10
aPaschalTable(0,15) = 3
aPaschalTable(1,15) = 30
aPaschalTable(0,16) = 4
aPaschalTable(1,16) = 17
aPaschalTable(0,17) = 4
aPaschalTable(1,17) = 7
aPaschalTable(0,18) = 3
aPaschalTable(1,18) = 27
FUNCTION paschal_moon( yy )
	DIM	gn	' golden number
	
	gn = yy MOD 19 + 1 - 1	'(-1) to make golden-number an index
	paschal_moon = jdateFromGregorian( aPaschalTable(1,gn), _
									aPaschalTable(0,gn), yy )

END FUNCTION



SUB easter_mallen( d, m, ByVal y )

	d = 0
	m = 0

	' Calculate Easter Sunday date
	Dim FirstDig, Remain19, temp	'intermediate results (all integers)
	Dim tA, tB, tC, tD, tE			'table A to E results (all integers)

	FirstDig = y \ 100				'first 2 digits of year (\ means integer division)
	Remain19 = y Mod 19				'remainder of year / 19

	' calculate PFM date
	tA = ((225 - 11 * Remain19) Mod 30) + 21

	' find the next Sunday
	tB = (tA - 19) Mod 7
	tC = (40 - FirstDig) Mod 7

	temp = y Mod 100
	tD = (temp + temp \ 4) Mod 7

	tE = ((20 - tB - tC - tD) Mod 7) + 1
	d = tA + tE

	'10 days were 'skipped' in the Gregorian calendar from 5-14 Oct 1582
	temp = 10
	'Only 1 in every 4 century years are leap years in the Gregorian
	'calendar (every century is a leap year in the Julian calendar)
	If 1600 < y Then temp = temp + FirstDig - 16 - ((FirstDig - 16) \ 4)
	d = d + temp

	' return the date
	If 61 < d Then
		d = d - 61
		m = 5       'for method 2, Easter Sunday can occur in May
	ElseIf 31 < d Then
		d = d - 31
		m = 4
	Else
		m = 3
	End If

END SUB


FUNCTION passover_conway( ByVal y )

	DIM	j
	
	j = roshhashanah_conway( y )
	
	DIM	dd,mm,yy
	
	gregorianFromJDate dd, mm, yy, j
	
	IF 10 = mm THEN dd = dd + 30
	
	passover_conway = jdateFromGregorian( 21, 3, yy ) + mm


END FUNCTION


FUNCTION roshhashanah_conway( ByVal y )

	DIM	g	' golden number
	
	g = y MOD 19 + 1
	
	DIM	n
	DIM	r
	DIM	x
	
	x = (12*g) MOD 19
	
	n = ( y\100 - y\400 - 2 ) _
			+	765433/492480 * x _
			+	(y MOD 4)/4 _
			-	(313 * y + 89081)/98496
	
	r = n - INT(n)
	
	DIM	j
	
	j = jdateFromGregorian( INT(n), 9, y )
	
	DIM	wd
	
	wd = weekdayFromJDate( j )	'1=sunday
	
	SELECT CASE wd
	CASE 1,4,6	'sun,wed,fri
		j = j + 1
	CASE 2		'mon
		IF 23269/25920 <= r  AND  11 < x THEN
			j = j + 1
		END IF
	CASE 3		'tue
		IF 1367/2160 <= r  AND 6 < x THEN
			j = j + 2
		END IF
	END SELECT
	
	roshhashanah_conway = j
	

END FUNCTION



PUBLIC aDateSplit
SUB parseLogDate( dd, mm, yy, sDate )

	aDateSplit = SPLIT(sDate,"-",3,vbTextCompare)
	IF 2 = UBound(aDateSplit) THEN
		yy = 0 + aDateSplit(0)
		mm = 0 + aDateSplit(1)
		dd = 0 + aDateSplit(2)
	ELSE
		yy = 0
		mm = 0
		dd = 0
	END IF

END SUB



SUB parseLogDateEx( dd, mm, yy, sDate )

	DIM	dDate

	IF 0 < Len(sDate) THEN
		CALL parseLogDate( dd, mm, yy, sDate )
	ELSE
		dDate = DATE
		dd = DAY( dDate )
		mm = MONTH( dDate )
		yy = YEAR( dDate )
	END IF

END SUB


SUB gregorianFromVBDate( dd, mm, yy, dDate )

	dd = DAY( dDate )
	mm = MONTH( dDate )
	yy = YEAR( dDate )

END SUB

'FUNCTION jdateFromVBDate( dDate )
'	DIM	dd, mm, yy
'	dd = DAY( dDate )
'	mm = MONTH( dDate )
'	yy = YEAR( dDate )
'	jdateFromVBDate = jdateFromGregorian( dd, mm, yy )
'END FUNCTION


FUNCTION jdateFromVBDate( dDate )
	DIM	dd, mm, yy
	CALL gregorianFromVBDate( dd, mm, yy, dDate )
	jdateFromVBDate = jdateFromGregorian( dd, mm, yy )
END FUNCTION







FUNCTION logDateFromDDMMYY( dd, mm, yy )
	DIM	sTemp

	sTemp = CSTR(yy) & " "
	IF mm < 10 THEN
		sTemp = sTemp & "0" & mm & " "
	ELSE
		sTemp = sTemp & mm & " "
	END IF
	IF dd < 10 THEN
		sTemp = sTemp & "0" & dd
	ELSE
		sTemp = sTemp & dd
	END IF
	logDateFromDDMMYY = sTemp
END FUNCTION


FUNCTION logDateFromJDate( jd )
	DIM	dd, mm, yy
	CALL gregorianFromJDate( dd, mm, yy, jd )
	logDateFromJDate = logDateFromDDMMYY( dd, mm, yy )
END FUNCTION


SUB parseLogTime( hh, mm, ss, sTime )

	aDateSplit = SPLIT(sTime,":",3,vbTextCompare)
	hh = 0 + aDateSplit(0)
	mm = 0 + aDateSplit(1)
	ss = 0 + aDateSplit(2)

END SUB













' 0=new, 1=1st quarter, 2=full, 3=last
FUNCTION moon_phaseFull( n, nPhase, nTZ, BYREF jd, BYREF frac )
	CONST PI = 3.14159265358979323846
	DIM	RAD
	RAD = PI/180.0
	DIM	i	' int
	DIM	am, az, c, t, t2, xtra
	DIM	jdate
	
	c = CDBL(n) + CDBL(nPhase) / 4.0
	t = c / 1236.85
	t2 = t * t
	az = 359.2242 + 29.10535608 * c
	am = 306.0253 + 385.81691806 * c + 0.0107306 * t2
	jdate = 2415020 + 28 * CLNG(n) + 7 * nPhase
	xtra = 0.75933 + 1.53058868 * c + ( CDBL(1.178e-4) - CDBL(1.55e-7) * t ) * t2
	SELECT CASE nPhase
	CASE 0, 2
		xtra = xtra + (( 0.1734 - CDBL(3.93e-4) * t ) * SIN( RAD * az ) - 0.4068 * sin( RAD * am ))
	CASE 1, 3
		xtra = xtra + (( 0.1721 - CDBL("4.0e-4") * t ) * sin( RAD * az ) - 0.6280 * sin( RAD * am ))
	END SELECT
	IF 0.0 <= xtra THEN
		i = CLNG(ROUND(xtra))
	ELSE
		i = CLNG(ROUND(xtra-1.0))
	END IF
	jd = jdate + FIX(i)
	frac = xtra - CDBL(i)
	moon_phaseFull = FIX(CDBL(jdate) + xtra + 0.5 - (CDBL(nTZ)/24.0))


END FUNCTION










%>