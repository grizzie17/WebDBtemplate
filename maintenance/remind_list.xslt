<?xml version='1.0'?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions">


<xsl:output method="html" indent="no" encoding="iso-8859-1"/>

<xsl:template match="/reminders">
	<xsl:choose>
		<xsl:when test="include">
			<xsl:apply-templates select="include" />
		</xsl:when>
		<xsl:otherwise>
			<table cellspacing="0" cellpadding="2" border="1">
			<tr>
			<th>Date</th>
			<th>Message</th>
			<th>Style</th>
			<th>Categories</th>
			</tr>
				<xsl:apply-templates select="event">
					<xsl:sort select="date/@type" order="descending"/>
					<xsl:sort select="date" order="ascending"/>
					<xsl:sort select="date/@offset" order="ascending"/>
					<xsl:sort select="subject"/>
				</xsl:apply-templates>
			</table>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template match="include">
	<b>Entries Imported from external source</b><br/>
	<xsl:value-of select="local:buildIncludeInfo(.)" />
</xsl:template>



<xsl:template match="event">
	<tr>
	<td nowrap="yes"><xsl:apply-templates select="date"/></td>
	<td><xsl:if test="style"><xsl:attribute name="class"><xsl:value-of select="style"/></xsl:attribute></xsl:if>
		<xsl:value-of select="subject"/>
		<xsl:apply-templates select="content"/>
		<xsl:apply-templates select="location"/>
		<xsl:apply-templates select="comments"/>
		</td>
	<td class="stylename">
		<xsl:apply-templates select="style" />
	</td>
	<td class="categoryname">
	<xsl:apply-templates select="category">
		<xsl:sort select="." order="ascending"/>
	</xsl:apply-templates>
	</td>
	</tr>
</xsl:template>

<xsl:template match="date">
	<xsl:value-of select="local:buildDateString(.)"/>
	<xsl:if test="@time"><br/>time= <xsl:value-of select="@time"/></xsl:if>
	<xsl:if test="@offset"><br/>offset= <xsl:value-of disable-output-escaping="yes" select="local:buildOffsetString(.)"/></xsl:if>
	<xsl:if test="@duration"><br/>duration= <xsl:value-of select="@duration"/></xsl:if>
	<xsl:if test="@pending"><br/><i>pending= <xsl:value-of select="@pending"/></i></xsl:if>
</xsl:template>

<xsl:template match="content">
	-- <i><xsl:value-of select="."/></i>
</xsl:template>

<xsl:template match="location">
	(<xsl:value-of select="."/>)
</xsl:template>

<xsl:template match="comments">
	<br/><xsl:value-of select="."/>
</xsl:template>

<xsl:template match="style">
<xsl:value-of select="."/>
</xsl:template>

<xsl:template match="category">
	<xsl:value-of select="."/><xsl:if test="position() != last()">, </xsl:if>
</xsl:template>


<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[


FUNCTION nameFromInterval( s )
	DIM	sName
	SELECT CASE LCASE( s )
	CASE "h"
		sName = "Hour"
	CASE "d"
		sName = "Day"
	CASE "w", "ww"
		sName = "Week"
	CASE "m"
		sName = "Month"
	CASE "y"
		sName = "Year"
	CASE ELSE
		sName = s
	END SELECT
	nameFromInterval = sName
END FUNCTION


FUNCTION buildIncludeInfo( o )

	DIM	oInc
	SET oInc = o.item(0)
	DIM	sHREF
	sHREF = oInc.getAttribute("href")
	
	DIM	i
	i = INSTR(sHREF, "?" )
	sHREF = MID(sHREF,i+1)
	DIM	a
	a = SPLIT(sHREF, "&")
	
	DIM	sOut
	DIM	aB
	DIM	q
	sOut = ""
	FOR EACH q IN a
		aB = SPLIT(q,"=")
		SELECT CASE LCASE(aB(0))
		CASE "h"
			sOut = sOut & "; Remote Server = " & aB(1)
		CASE "q"
			sOut = sOut & "; Calendar Source = " & aB(1)
		CASE "i"
			sOut = sOut & "; Refresh Interval = " & aB(1)
		CASE "v"
			sOut = sOut & "; Refresh Amount = " & aB(1)
		CASE "b"
			sOut = sOut & "; Force Refresh On = " & aB(1)
		CASE ELSE
			sOut = sOut & "; (" & q & ")"
		END SELECT
	NEXT 'q
	
	DIM	sCacheInt
	DIM	sCacheVal
	DIM	sCacheBrk
	sCacheInt = oInc.getAttribute("cache-interval")
	sCacheVal = oInc.getAttribute("cache-value")
	sCacheBrk = oInc.getAttribute("cache-break-interval")
	IF NOT ISNULL(sCacheInt)  AND  NOT ISNULL(sCacheVal) THEN
		IF ISNUMERIC(sCacheVal) THEN
			sOut = sOut & "; Refresh Interval = " & sCacheVal _
				&	" " & nameFromInterval( sCacheInt )
			IF 1 < CINT(sCacheVal) THEN
				sOut = sOut & "s"
			END IF
			IF NOT ISNULL(sCacheBrk) THEN
				sOut = sOut & "; Force Refresh On = " & nameFromInterval( sCacheBrk)
			END IF
		END IF
	END IF
	
	sOut = MID(sOut,2)
	buildIncludeInfo = sOut
	
	SET oInc = Nothing

END FUNCTION





FUNCTION formatWeeknumber( wn )
	DIM	s

	SELECT CASE wn
	CASE 1
		s = "First"
	CASE 2
		s = "Second"
	CASE 3
		s = "Third"
	CASE 4
		s = "Fourth"
	CASE 5
		s = "Fifth"
	CASE -1
		s = "Last"
	CASE -2
		s = "Next to Last"
	CASE -3
		s = "Third from Last"
	CASE ELSE
		s = CSTR(wn)
	END SELECT
	formatWeeknumber = s
END FUNCTION


FUNCTION formatMonthly( cc, cm )
	DIM	s
	SELECT CASE cc
	CASE 1
		s = "Every Month"
	CASE 2
		IF cm MOD 2 THEN
			s = "Odd"
		ELSE
			s = "Even"
		END IF
		s = s & " Bimonthly"
	CASE 3
		s = "Quarterly(" & cm & ")"
	CASE 6
		s = "Semiannually(" & cm & ")"
	CASE ELSE
		s = "?"
	END SELECT
	formatMonthly = s
END FUNCTION



FUNCTION dateString_single( oDate )
	DIM	aSplitDate
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	dateString_single = aSplitDate(2) & "-" & MONTHNAME(CINT(aSplitDate(1)),TRUE) & "-" & aSplitDate(0)
	'dateString_single = oDate.text
END FUNCTION



FUNCTION dateString_weekly( oDate )
	DIM	aSplitDate
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	DIM	sRecurr
	SELECT CASE aSplitDate(0)
	CASE "0"
		sRecurr = "Every"
	CASE "1"
		sRecurr = "Odd"
	CASE "2"
		sRecurr = "Even"
	CASE ELSE
		sRecurr = aSplitDate(0)
	END SELECT
	DIM	aDays
	aDays = SPLIT( aSplitDate(1), ",", -1, vbTextCompare )
	DIM	s
	DIM	n
	s = ""
	FOR EACH n IN aDays
		IF n < 1 THEN n = 1
		IF 7 < n THEN n = 7
		s = s & ", " & WEEKDAYNAME(CINT(n))
	NEXT
	s = MID(s, 2)
	dateString_weekly = sRecurr & " (" & s & ")"
END FUNCTION




FUNCTION dateString_monthlyDayN( oDate )

	DIM	cc, cm, dd
	DIM	s
	DIM	aSplitDate

	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	cc = CINT(aSplitDate(0))
	cm = CINT(aSplitDate(1))
	dd = CINT(aSplitDate(2))
	IF dd < 0 THEN
		s = "end(" & CSTR(-dd) & ")"
	ELSE
		s = CSTR(dd)
	END IF
	dateString_monthlyDayN = s & " " & formatMonthly( cc, cm )
END FUNCTION

FUNCTION dateString_monthlyWDay( oDate )
	DIM	cc, cm, wn, wd
	DIM	swd
	DIM	aSplitDate
	
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	cc = CINT(aSplitDate(0))
	cm = CINT(aSplitDate(1))
	wn = CINT(aSplitDate(2))
	wd = CINT(aSplitDate(3))
	SELECT CASE wd
	CASE -1
		swd = "weekday"
	CASE -2
		swd = "weekend"
	CASE ELSE
		swd = WEEKDAYNAME(wd,TRUE,1)
	END SELECT
	dateString_monthlyWDay = formatWeeknumber(wn) & " " & swd & " in " & formatMonthly( cc, cm )
END FUNCTION

FUNCTION dateString_monthly( oDate )
	SELECT CASE oDate.getAttribute("monthly")
	CASE "dayn"
		dateString_monthly = dateString_monthlyDayN( oDate )
	CASE "wday"
		dateString_monthly = dateString_monthlyWDay( oDate )
	CASE ELSE
		dateString_monthly = oDate.text
	END SELECT
END FUNCTION



FUNCTION dateString_yearlyDayN( oDate )
	'dateString_yearlyDayN = oDate.text
	DIM	aSplitDate
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	dateString_yearlyDayN = aSplitDate(1) & "-" & MONTHNAME(CINT(aSplitDate(0)),TRUE) & "-Yearly"
END FUNCTION

FUNCTION dateString_yearlyWDay( oDate )

	DIM	mm, wn, wd, swd
	DIM	aSplitDate

	'dateString_yearlyWDay = oDate.text
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	mm = CINT(aSplitDate(0))
	wn = CINT(aSplitDate(1))
	wd = CINT(aSplitDate(2))
	SELECT CASE wd
	CASE -1
		swd = "weekday"
	CASE -2
		swd = "weekend"
	CASE ELSE
		swd = WEEKDAYNAME(wd,TRUE,1)
	END SELECT

	dateString_yearlyWDay = formatWeeknumber(wn) & " " & swd & " in " & MONTHNAME(mm,TRUE)

END FUNCTION

FUNCTION dateString_yearly( oDate )
	SELECT CASE oDate.getAttribute("yearly")
	CASE "dayn"
		dateString_yearly = dateString_yearlyDayn( oDate )
	CASE "wday"
		dateString_yearly = dateString_yearlyWDay( oDate )
	CASE ELSE
		dateString_yearly = oDate.text
	END SELECT
END FUNCTION

FUNCTION dateString_keyword( oDate )
	dateString_keyword = oDate.text
END FUNCTION

FUNCTION dateString_season( oDate )
	SELECT CASE oDate.text
	CASE 0
		dateString_season = "Equinox (March)"
	CASE 1
		dateString_season = "Solstice (June)"
	CASE 2
		dateString_season = "Equinox (September)"
	CASE 3
		dateString_season = "Solstice (December)"
	END SELECT
END FUNCTION



FUNCTION buildDateString( o )

	DIM	oDate
	SET oDate = o.item(0)

	SELECT CASE oDate.getAttribute("type")
	CASE "single"
		buildDateString = dateString_single( oDate )
	CASE "weekly"
		buildDateString = dateString_weekly( oDate )
	CASE "monthly"
		buildDateString = dateString_monthly( oDate )
	CASE "yearly"
		buildDateString = dateString_yearly( oDate )
	CASE "keyword"
		buildDateString = dateString_keyword( oDate )
	CASE "season"
		buildDateString = dateString_season( oDate )
	CASE ELSE
		buildDateString = oDate.text
	END SELECT
	SET oDate = Nothing
END FUNCTION





FUNCTION offset_eval( aOffset )
	DIM	i
	DIM	s
	DIM	n
	DIM	t
	DIM	sPrev
	
	i = 0
	s = ""
	sPrev = ""
	DO WHILE i < UBOUND(aOffset)
		IF 0 < LEN(sPrev) THEN
			IF "?" <> sPrev THEN s = s & ";<br/>&nbsp;&nbsp;&nbsp;&nbsp;"
		END IF
		sPrev = aOffset(i)
		SELECT CASE aOffset(i)
		CASE "?"
			s = s & " if "
			IF ISNUMERIC( aOffset(i+1) ) THEN
				s = s & " " & WEEKDAYNAME( CINT(aOffset(i+1)), TRUE, 1 )
			ELSE
				s = s & " " & aOffset(i+1)
			END IF
			s = s & " then"
		CASE "+", "++", "-", "--", "~"
			IF ISNUMERIC( aOffset(i+1) ) THEN
				t = WEEKDAYNAME( CINT(aOffset(i+1)), TRUE, 1 )
			ELSE
				t = aOffset(i+1)
			END IF
			SELECT CASE aOffset(i)
			CASE "+"
				s = s & " " & t & " on or after"
			CASE "++"
				s = s & " " & t & " after"
			CASE "-"
				s = s & " " & t & " on or before"
			CASE "--"
				s = s & " " & t & " before"
			CASE "~"
				s = s & " nearest " & t
			END SELECT
		CASE "#"
			n = CINT(aOffset(i+1))
			IF n < 0 THEN
				s = s & " subtract " & CSTR(-n)
			ELSE
				s = s & " add " & CSTR(n)
			END IF
		END SELECT
		i = i + 2
	LOOP
	offset_eval = s
END FUNCTION



FUNCTION buildOffsetString( o )
	DIM oDate
	DIM	sOffset
	DIM	aSplitDate
	SET oDate = o.item(0)
	sOffset = oDate.getAttribute("offset")
	aSplitDate = SPLIT( sOffset, " ", -1, vbTextCompare )
	buildOffsetString = offset_eval( aSplitDate )
'	buildOffsetString = sOffset
	SET oDate = Nothing
END FUNCTION


]]></msxsl:script>
</xsl:stylesheet>
