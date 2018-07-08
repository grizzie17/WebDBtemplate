<?xml version="1.0"?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions">


<xsl:output method="html" indent="no" encoding="iso-8859-1"/>




<xsl:template match="/calendar">

		<xsl:apply-templates select="event">
			<xsl:sort select="category[substring(.,1,3) = 'al-']" />
			<xsl:sort select="date"/>
		</xsl:apply-templates>
</xsl:template>

<xsl:template match="event">
<xsl:if test="local:welcomeIn(.)">
<div class="neighbormeeting">
		<span style="font-weight:bold"><xsl:apply-templates select="subject"/></span><br/>
		<xsl:apply-templates select="content"/>
		<xsl:apply-templates select="body"/>
		<xsl:apply-templates select="location"/>
		<xsl:apply-templates select="comments"/>
	<br/><i>[Next Meeting: <xsl:value-of select="local:buildDateString(.)"/>]</i>
</div>
</xsl:if>
</xsl:template>


<xsl:template match="subject">
		<xsl:value-of disable-output-escaping="yes" select="." />
</xsl:template>

<xsl:template match="content">
 -- <i><xsl:value-of disable-output-escaping="yes" select="." /></i>
</xsl:template>

<xsl:template match="body">
	-- <i><xsl:value-of disable-output-escaping="yes" select="." /></i>
</xsl:template>

<xsl:template match="location">
	(<xsl:value-of disable-output-escaping="yes" select="." />)
</xsl:template>

<xsl:template match="comments">
	<br/><xsl:value-of disable-output-escaping="yes" select="." />
</xsl:template>

<xsl:template match="text()">
	<xsl:value-of select="."/>
</xsl:template>


<xsl:template match="small">
<span style="font-size: small"><xsl:apply-templates/></span>
</xsl:template>

<xsl:template match="large">
<span style="font-size: large"><xsl:apply-templates/></span>
</xsl:template>


<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[



DIM	g_bFirstRun
g_bFirstRun = true


FUNCTION isFirstRun()
	isFirstRun = g_bFirstRun
	g_bFirstRun = false
END FUNCTION


DIM	aNeighbors(26)
DIM	nNeighbors
nNeighbors = 0

FUNCTION welcomeIn( oEvent )

	welcomeIn = TRUE

	DIM	s
	DIM	oCat
	SET oCat = oEvent.item(0).selectSingleNode("./category[substring(.,1,3) = 'al-']")
	s = oCat.Text
	SET oCat = Nothing
	DIM	bFound
	bFound = FALSE
	DIM	i
	FOR i = 0 to nNeighbors
		IF s = aNeighbors(i) THEN
			bFound = TRUE
			welcomeIn = FALSE
			EXIT FOR
		END IF
	NEXT
	IF NOT bFound THEN
		aNeighbors(nNeighbors) = s
		nNeighbors = nNeighbors + 1
	END IF

END FUNCTION









function buildDateString(oEvent)

	dim oAttrs
	dim dd
	dim	mm
	dim	yy
	dim wd
	dim oDate

	set oDate = oEvent.item(0).selectSingleNode("date")
	if not oDate is nothing then

		set oAttrs = oDate.attributes
		dd = oAttrs.getNamedItem("d").value
		mm = oAttrs.getNamedItem("m").value
		yy = oAttrs.getNamedItem("y").value
		wd = oAttrs.getNamedItem("wd").value
		
		buildDateString = WEEKDAYNAME(wd,false,1) & ", " & MONTHNAME(mm,0) & " " & dd & ", " & yy
		
		SET oAttrs = Nothing
		SET oDate = Nothing

	else

		buildDateString = ""
	end if
end function





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


FUNCTION buildDateStringDef( o )


	DIM oDate
	SET oDate = o.item(0).selectSingleNode("date")

	SELECT CASE oDate.getAttribute("type")
	CASE "single"
		buildDateStringDef = dateString_single( oDate )
	CASE "monthly"
		buildDateStringDef = dateString_monthly( oDate )
	CASE "yearly"
		buildDateStringDef = dateString_yearly( oDate )
	CASE "keyword"
		buildDateStringDef = dateString_keyword( oDate )
	CASE "season"
		buildDateStringDef = dateString_season( oDate )
	CASE ELSE
		buildDateStringDef = oDate.text
	END SELECT
	SET oDate = Nothing
END FUNCTION




]]></msxsl:script>
</xsl:stylesheet>