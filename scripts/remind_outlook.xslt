<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions">


<xsl:output method="text" media-type="text" indent="no" encoding="windows-1252"/>




<xsl:template match="/calendar">
Subject<xsl:call-template name="tab"/>Start Date<xsl:call-template name="tab"/>Start Time<xsl:call-template name="tab"/>End Date<xsl:call-template name="tab"/>End Time<xsl:call-template name="tab"/>All day event<xsl:call-template name="tab"/>Reminder on/off<xsl:call-template name="tab"/>Reminder Date<xsl:call-template name="tab"/>Reminder Time<xsl:call-template name="tab"/>Meeting Organizer<xsl:call-template name="tab"/>Required Attendees<xsl:call-template name="tab"/>Optional Attendees<xsl:call-template name="tab"/>Meeting Resources<xsl:call-template name="tab"/>Billing Information<xsl:call-template name="tab"/>Categories<xsl:call-template name="tab"/>Description<xsl:call-template name="tab"/>Location<xsl:call-template name="tab"/>Mileage<xsl:call-template name="tab"/>Priority<xsl:call-template name="tab"/>Private<xsl:call-template name="tab"/>Sensitivity<xsl:call-template name="tab"/>Show time as
<xsl:apply-templates select="event"><xsl:sort select="date"/><xsl:sort select="date/@time"/></xsl:apply-templates>
</xsl:template>

<xsl:template match="event">
<xsl:apply-templates select="subject" />
<xsl:apply-templates select="date"/><xsl:apply-templates select="style"/><xsl:apply-templates select="subject"/><xsl:apply-templates select="content"/><xsl:apply-templates select="location"/><xsl:apply-templates select="comments"/><xsl:if test="style">}}</xsl:if>|<xsl:call-template name="br" /></xsl:template>

<xsl:template match="date">
<xsl:value-of select="@m"/>/<xsl:value-of select="@d"/>/<xsl:value-of select="@y"/>
</xsl:template>

<xsl:template match="subject"><xsl:value-of disable-output-escaping="no" select="." /></xsl:template>


<xsl:template match="style">{{~~<xsl:value-of select="."/>~~:</xsl:template>


<xsl:template match="content"> -- {{i:<xsl:value-of disable-output-escaping="no" select="." />}}</xsl:template>

<xsl:template match="location"> (<xsl:value-of disable-output-escaping="no" select="." />)</xsl:template>

<xsl:template match="comments"><xsl:call-template name="br"/><xsl:value-of disable-output-escaping="no" select="." /></xsl:template>

<xsl:template match="text()"><xsl:value-of select="."/></xsl:template>


<xsl:template match="small">
{{small:<xsl:apply-templates/>}}</xsl:template>

<xsl:template match="large">
{{big:<xsl:apply-templates/>}}</xsl:template>


<xsl:template name="tab"><xsl:text>&#7;</xsl:text></xsl:template>
<xsl:template name="br"><xsl:text>&#13;&#10;</xsl:text></xsl:template>


<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[


DIM g_sLastDay
DIM g_nLastMonth
DIM g_nNextSat
DIM g_bNewDay

g_sLastDay = ""
g_nLastMonth = 0
g_nNextSat = 0
g_bNewDay = false






function encodeText(oNode)
	DIM s
	s = oNode.item(0).text
	''s = REPLACE( s, "<", "&lt;", 1, -1, vbTextCompare )
	''s = REPLACE( s, ">", "&gt;", 1, -1, vbTextCompare )
	''s = REPLACE( s, "&", "&amp;", 1, -1, vbTextCompare )
	encodeText = s
end function


function getMonthName(m,n)
	getMonthName = MONTHNAME( m, n )
end function


function getDayName(d,n)
	getDayName = WEEKDAYNAME( d, TRUE, 1 )
end function

function isNewDay(oDate)

	DIM sTemp

	sTemp = oDate.text

	if sTemp = g_sLastDay then
		isNewDay = false
	else
		g_sLastDay = sTemp
		isNewDay = true
	end if
end function

sub calcNextSat(oDate)

	DIM dd, wd
	DIM nDay, nWd
	DIM oAttrs

	SET oAttrs = oDate.item(0).attributes
	dd = oAttrs.getNamedItem("d").value
	wd = oAttrs.getNamedItem("wd").value
	nDay = CINT(dd)
	nWd = CINT(wd)
	g_nNextSat = (7 - nWd) + nDay
end sub


function isNewWeek(oDate)
	dim oAttrs
	dim dd, nDay

	SET oAttrs = oDate.item(0).attributes
	dd = oAttrs.getNamedItem("d").value
	nDay = CINT(dd)

	if g_nNextSat < nDay then
		calcNextSat oDate
		isNewWeek = true
	else
		isNewWeek = false
	end if

end function

function isNewMonth(oDate)
	dim oAttrs
	dim mm

	SET oAttrs = oDate.item(0).attributes
	mm = oAttrs.getNamedItem("m").value

	if mm = g_nLastMonth then
		isNewMonth = false
	else
		calcNextSat oDate

		g_nLastMonth = mm
		isNewMonth = true
	end if
end function

function buildDateString(oEvent)

	dim oAttrs
	dim dd
	dim oDate

	set oDate = oEvent.item(0).selectSingleNode("date")
	if isNewDay( oDate ) then

		set oAttrs = oDate.attributes
		dd = oAttrs.getNamedItem("d").value

		g_bNewDay = true
		buildDateString = "" & dd

	else

		g_bNewDay = false
		buildDateString = ""
	end if
end function

function buildWeekdayString(oEvent)
	dim oAttrs
	dim wd
	dim oDate

	if g_bNewDay then
		set oDate = oEvent.item(0).selectSingleNode("date")
		set oAttrs = oDate.attributes
		wd = oAttrs.getNamedItem("wd").value
	
		g_bNewDay = false
		buildWeekdayString = getDayName(wd,3)
	else
		buildWeekdayString = ""
	end if
end function

function buildMonthString(oDate)
	dim oAttrs
	dim mm, yy

	SET oAttrs = oDate.item(0).attributes
	mm = oAttrs.getNamedItem("m").value
	yy = oAttrs.getNamedItem("y").value

	buildMonthString = getMonthName(mm,0) & " - " & yy
end function



]]></msxsl:script>
</xsl:stylesheet>