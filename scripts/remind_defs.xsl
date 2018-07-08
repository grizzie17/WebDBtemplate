<?xml version='1.0'?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl" language="VBScript">

<xsl:template match="/">
	<table cellspacing="0" cellpadding="2" border="1">
	<tr>
	<th>Date</th>
	<th>Message</th>
	</tr>
		<xsl:for-each select="//event" order-by="-date/@type;+date;+date/@offset;+subject">
			<xsl:apply-templates select="."/>
		</xsl:for-each>
	</table>
</xsl:template>


<xsl:template match="event">
	<tr>
	<td nowrap="yes"><xsl:apply-templates select="date"/></td>
	<td><xsl:if test="style"><xsl:attribute name="class"><xsl:value-of select="style"/></xsl:attribute></xsl:if>
		<xsl:value-of select="subject"/>
		<xsl:apply-templates select="body"/>
		<xsl:apply-templates select="location"/>
		<xsl:apply-templates select="comments"/>
		</td>
	</tr>
</xsl:template>

<xsl:template match="date">
	<xsl:eval>buildDateString(me)</xsl:eval>
	<xsl:if test="@offset"><br/>offset= <xsl:eval>buildOffsetString(me)</xsl:eval></xsl:if>
	<xsl:if test="@duration"><br/>duration= <xsl:value-of select="@duration"/></xsl:if>
</xsl:template>

<xsl:template match="body">
	-- <i><xsl:value-of select="."/></i>
</xsl:template>

<xsl:template match="location">
	(<xsl:value-of select="."/>)
</xsl:template>

<xsl:template match="comments">
	<br/><xsl:value-of select="." />
</xsl:template>


<xsl:template match="category">
	<xsl:value-of /><xsl:if test="context()[not(end())]">, </xsl:if>
</xsl:template>


<xsl:script><![CDATA[

PUBLIC aSplitDate


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
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	dateString_single = aSplitDate(2) & "-" & MONTHNAME(CINT(aSplitDate(1)),TRUE) & "-" & aSplitDate(0)
	'dateString_single = oDate.text
END FUNCTION




FUNCTION dateString_monthlyDayN( oDate )

	DIM	cc, cm, dd
	DIM	s

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
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	cc = CINT(aSplitDate(0))
	cm = CINT(aSplitDate(1))
	wn = CINT(aSplitDate(2))
	wd = CINT(aSplitDate(3))
	dateString_monthlyWDay = formatWeeknumber(wn) & " " & WEEKDAYNAME(wd,TRUE,1) & " in " & formatMonthly( cc, cm )
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
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	dateString_yearlyDayN = aSplitDate(1) & "-" & MONTHNAME(CINT(aSplitDate(0)),TRUE) & "-Yearly"
END FUNCTION

FUNCTION dateString_yearlyWDay( oDate )

	DIM	mm, wn, wd

	'dateString_yearlyWDay = oDate.text
	aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
	mm = CINT(aSplitDate(0))
	wn = CINT(aSplitDate(1))
	wd = CINT(aSplitDate(2))

	dateString_yearlyWDay = formatWeeknumber(wn) & " " & WEEKDAYNAME(wd,TRUE,1) & " in " & MONTHNAME(mm,TRUE)

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



FUNCTION buildDateString( oDate )

	SELECT CASE oDate.getAttribute("type")
	CASE "single"
		buildDateString = dateString_single( oDate )
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
			IF "?" <> sPrev THEN s = s & ";<br>" & vbCRLF
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



FUNCTION buildOffsetString( oDate )
	DIM	sOffset
	sOffset = oDate.getAttribute("offset")
	aSplitDate = SPLIT( sOffset, " ", -1, vbTextCompare )
	buildOffsetString = offset_eval( aSplitDate )
'	buildOffsetString = sOffset
END FUNCTION


]]></xsl:script>
</xsl:stylesheet>
