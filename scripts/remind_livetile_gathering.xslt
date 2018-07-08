<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions">


	<xsl:output method="text" media-type="text" indent="no" encoding="windows-1252"/>



	<xsl:template match="/calendar">

		<xsl:apply-templates select="event">
			<xsl:sort select="date"/>
			<xsl:sort select="date/@time" order="ascending"/>
		</xsl:apply-templates>
	</xsl:template>

<xsl:template match="event">
<xsl:if test="local:isFirstRun()"><xsl:value-of select="local:buildDateString(.)"/> - <xsl:apply-templates select="subject"/><xsl:apply-templates select="content"/><xsl:apply-templates select="body"/><xsl:apply-templates select="location"/><xsl:apply-templates select="comments"/></xsl:if>
</xsl:template>


	<xsl:template match="subject">
		<xsl:value-of disable-output-escaping="yes" select="." />
	</xsl:template>

	<xsl:template match="content"> -- <xsl:value-of disable-output-escaping="yes" select="." /></xsl:template>

	<xsl:template match="body"> -- <xsl:value-of disable-output-escaping="yes" select="." /></xsl:template>

	<xsl:template match="location"> (<xsl:value-of disable-output-escaping="yes" select="." />)</xsl:template>

	<xsl:template match="comments">
		<br/>
		<xsl:value-of disable-output-escaping="yes" select="." />
	</xsl:template>

	<xsl:template match="text()"><xsl:value-of select="."/></xsl:template>


	<xsl:template match="small">
		<span style="font-size: small">
			<xsl:apply-templates/>
		</span>
	</xsl:template>

	<xsl:template match="large">
		<span style="font-size: large">
			<xsl:apply-templates/>
		</span>
	</xsl:template>


	<msxsl:script implements-prefix="local" language="VBscript">
		<![CDATA[



DIM	g_bFirstRun
g_bFirstRun = true


FUNCTION isFirstRun()
	isFirstRun = g_bFirstRun
	g_bFirstRun = false
END FUNCTION






function getMonthName(m,n)
	getMonthName = MONTHNAME( m, n )
end function







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
		
		buildDateString = LEFT(WEEKDAYNAME(wd,false,1),3) & " " & mm & "/" & dd & "/" & yy

	else

		buildDateString = ""
	end if
end function




]]></msxsl:script>
</xsl:stylesheet>