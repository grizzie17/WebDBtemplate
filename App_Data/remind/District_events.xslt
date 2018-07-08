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
<category>al-a</category>
<category>al-b</category>
<category>al-c</category>
<category>al-d</category>
<category>al-e</category>
<category>al-f</category>
<category>al-g</category>
<category>al-h</category>
<category>al-i</category>
<category>al-j</category>
<category>al-k</category>
<category>al-l</category>
<category>al-m</category>
<category>al-n</category>
<category>al-o</category>
<category>al-p</category>
<category>al-q</category>
<category>al-r</category>
<category>al-s</category>
<category>al-t</category>
<category>al-u</category>
<category>al-v</category>
<category>al-w</category>
<category>al-x</category>
<category>al-y</category>
<category>al-z</category>
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="subject">
<subject><xsl:value-of disable-output-escaping="yes" select="local:replaceLocalURI(.)" /></subject>
</xsl:template>
file:///C:/Documents and Settings/John/My Documents/WebWork/Inetpub/wwwroot/dbDrivenTemplate/App_Data/remind/External_events;motorcycle;categories.xml
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