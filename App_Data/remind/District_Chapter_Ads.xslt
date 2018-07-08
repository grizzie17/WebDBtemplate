<?xml version="1.0"?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions"
>

<xsl:output method="xml" version="1.0" indent="yes" encoding="ISO-8859-1" />


<xsl:template match="/|@*|node()">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="@time" />



<xsl:template match="style[.='RmdBusinessMeeting']">
<style>RmdDistrictBusinessMeeting</style>
</xsl:template>

<xsl:template match="style[.='RmdChapterActivity']">
<style>RmdDistrictBusinessMeeting</style>
</xsl:template>

<xsl:template match="style[.='RmdSafety']">
<style>RmdDistrictEvent</style>
</xsl:template>

<xsl:template match="style[.='RmdSafetySup']">
<style>RmdSafetyEnc</style>
</xsl:template>

<xsl:template match="style[.='RmdSafetyEnc']">
<style>RmdSafetyInfo</style>
</xsl:template>

<xsl:template match="category[.='districtrally']">
<category>al-b</category>
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="subject">
<subject><xsl:value-of disable-output-escaping="yes" select="local:replaceLocalURI(.)" /></subject>
</xsl:template>

<xsl:template match="content">
<content><xsl:value-of disable-output-escaping="yes" select="local:replaceLocalURI(.)" /></content>
</xsl:template>

<xsl:template match="location">
<location><xsl:value-of disable-output-escaping="yes" select="local:replaceLocalURI(.)" /></location>
</xsl:template>

<xsl:template match="comments">
<location><xsl:value-of disable-output-escaping="yes" select="local:replaceLocalURI(.)" /></location>
</xsl:template>

<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[

FUNCTION replaceLocalURI( o )
	DIM	s
	s = o.item(0).text
	IF 0 < LEN(s) THEN
		s = REPLACE(s, "local:", "http://www.Alabama-gwrra.org/")
		
		DIM	oReg
		DIM	oMatchList
		DIM	oMatch
		SET oReg = NEW RegExp
		IF NOT oReg IS Nothing THEN
			oReg.Pattern = "\{\{(picture|image|img|) ([\w\.]+)( [\w%=]+)*\}\}"
			oReg.IgnoreCase = TRUE
			oReg.Global = TRUE
			s = oReg.Replace( s, "" )
		END IF
	END IF
	replaceLocalURI = "<" & "!" & "[CDATA[" & s & "]" & "]" & ">"
END FUNCTION

]]></msxsl:script>

</xsl:stylesheet>