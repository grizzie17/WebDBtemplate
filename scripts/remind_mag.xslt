<?xml version='1.0'?>
<xsl:stylesheet
	version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	>


<xsl:output method="xml" version="1.0" indent="yes" encoding="iso-8859-1"/>


<xsl:template match="/">
		<xsl:apply-templates>
			<xsl:sort select="date"/>
			<xsl:sort select="date/@time" order="ascending"/>
		</xsl:apply-templates>
</xsl:template>

<xsl:template match="node()|@*">
	<xsl:copy>
		<xsl:apply-templates select="node()|@*">
			<xsl:sort select="date"/>
			<xsl:sort select="date/@time" order="ascending"/>
		</xsl:apply-templates>
	</xsl:copy>
</xsl:template>




</xsl:stylesheet>