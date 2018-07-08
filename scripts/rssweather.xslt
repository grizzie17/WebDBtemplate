<?xml version="1.0"?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:content="http://purl.org/rss/1.0/modules/content/"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:yweather="http://xml.weather.yahoo.com/ns/rss/1.0"
	xmlns:local="#local-functions"
	>

<xsl:output method="html" indent="no" encoding="ISO-8859-1"/>


<xsl:template match="/rss">
		<xsl:apply-templates select="channel"/>
</xsl:template>

<xsl:template match="channel">
	<!--div class="rssTitle">
	<xsl:apply-templates select="image"/><br/>
	<b><a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute><xsl:value-of disable-output-escaping="yes" select="title"/></a></b>
	</div-->
	<!--xsl:apply-templates select="description"/-->
	<!--xsl:choose>
		<xsl:when test="pubDate">
	<xsl:apply-templates select="pubDate"/>
		</xsl:when>
		<xsl:otherwise>
	<xsl:apply-templates select="lastBuildDate"/>
		</xsl:otherwise>
	</xsl:choose-->
	
	<xsl:apply-templates select="item"/>
	<xsl:apply-templates select="yweather:location" />
</xsl:template>

<xsl:template match="yweather:location">
<div style="color: #999999; font-size: smaller; font-family: sans-serif;">
Location: <xsl:value-of select="@city"/>, <xsl:value-of select="@region"/>
</div>
</xsl:template>

<xsl:template match="yweather:condition">
	<div>
	<img height="26" width="26" align="bottom"><xsl:attribute name="alt"><xsl:value-of select="@text"/></xsl:attribute><xsl:attribute name="src">images/weather/<xsl:value-of select="@code"/>.gif</xsl:attribute></img>
	<span class="weatherCurrentTemp"><xsl:value-of select="@temp"/><xsl:text disable-output-escaping="yes">_$$deg;</xsl:text></span>
	<xsl:text> </xsl:text>
	<span class="weatherCurrentText"><xsl:value-of select="@text"/></span>
	<xsl:if test="@temp != ../yweather:wind/@chill">something</xsl:if>
	</div>
</xsl:template>

<xsl:template match="yweather:forecast">
	<div>
	<span class="weatherForecastDay"><xsl:value-of select="@day"/></span>
	<xsl:text> </xsl:text>
	<img height="26" width="26" align="bottom"><xsl:attribute name="alt"><xsl:value-of select="@text"/></xsl:attribute><xsl:attribute name="src">images/weather/<xsl:value-of select="@code"/>.gif</xsl:attribute></img>
	<span class="weatherForecastText"><xsl:value-of select="@text"/></span>
	<xsl:text> </xsl:text>
	<span class="weatherForecastHigh"><xsl:value-of select="@high"/><xsl:if test="@high != '-'"><xsl:text disable-output-escaping="yes">_$$deg;</xsl:text></xsl:if></span>_$$nbsp;<span class="weatherForecastLow"><xsl:value-of select="@low"/><xsl:if test="@low != '-'"><xsl:text disable-output-escaping="yes">_$$deg;</xsl:text></xsl:if></span>
	</div>
</xsl:template>

<xsl:template match="pubDate">
	<xsl:if test="local:uniqueDate(.)">
	<div style="color: #999999; font-size: xx-small; font-family: sans-serif;">
	Updated: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
	</xsl:if>
</xsl:template>

<xsl:template match="lastBuildDate">
	<div style="color: #999999; font-size: smaller; font-family: sans-serif;">
	Updated: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
</xsl:template>

<xsl:template match="description">
	<div>
	<xsl:value-of disable-output-escaping="yes" select="."/>
	</div>
</xsl:template>

<xsl:template match="image">
	<a><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute>
	<img border="0" align="absbottom"><xsl:attribute name="src"><xsl:value-of select="url"/></xsl:attribute></img></a>
</xsl:template>

<xsl:template match="item">
	<xsl:apply-templates select="yweather:condition"/>
	<xsl:apply-templates select="yweather:forecast"/>
	<!--div class="rssItem">
	<div class="rssItemTitle">
	<xsl:choose>
	<xsl:when test="link">
		<b><a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute><xsl:value-of disable-output-escaping="yes" select="title"/></a></b>
	</xsl:when>
	<xsl:otherwise>
		<b><xsl:value-of disable-output-escaping="yes" select="title"/></b>
	</xsl:otherwise>
	</xsl:choose>
	</div>
	
	<xsl:if test="title != description  and  link != description">
	<div class="rssItemDescription">
	<xsl:value-of disable-output-escaping="yes" select="description"/>
	</div>
	</xsl:if>
	<xsl:apply-templates select="content:encoded"/>
	<xsl:apply-templates select="enclosure"/>
	</div-->
	<xsl:apply-templates select="pubDate"/>
</xsl:template>

<xsl:template match="content:encoded">
	<div class="rssContent">
	<xsl:value-of disable-output-escaping="yes" select="."/>
	</div>
</xsl:template>

<xsl:template match="enclosure">
	<div class="rssEnclosure">
	<img ><xsl:attribute name="src"><xsl:value-of select="@url"/></xsl:attribute></img>
	</div>
</xsl:template>


<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[




function encodeText(oNode)
	DIM s
	s = oNode.item(0).text
	''s = REPLACE( s, "<", "&lt;", 1, -1, vbTextCompare )
	''s = REPLACE( s, ">", "&gt;", 1, -1, vbTextCompare )
	''s = REPLACE( s, "&", "_$$", 1, -1, vbTextCompare )
	encodeText = s
end function


DIM g_sPrevDate
g_sPrevDate = ""

FUNCTION uniqueDate( oNode )
	DIM	s
	s = oNode.item(0).text
	IF s <> g_sPrevDate THEN
		uniqueDate = TRUE
	ELSE
		uniqueDate = FALSE
	END IF
	g_sPrevDate = s
END FUNCTION

FUNCTION encodeDate( oNode )
	DIM s
	DIM	d
	s = oNode.item(0).text
	encodeDate = s
	EXIT FUNCTION
	IF ISDATE( s ) THEN
		d = CDATE( s )
		encodeDate = FORMATDATETIME( d, 1 )
	ELSE
		DIM	i
		i = INSTR( s, "," )
		IF 0 < i THEN
			DIM	t
			t = TRIM(MID( s, i+1 ))
			IF ISDATE( t ) THEN
				d = CDATE( t )
				encodeDate = CSTR(FORMATDATETIME( d, 1 ))
			ELSE
				DIM	a
				a = SPLIT( t, " " )
				t = a(0) & "-" & a(1) & "-" & a(2)
				FOR i = 3 TO UBOUND(a)
					t = t & " " & a(i)
				NEXT 'i
				'd = CDATE( t )
				encodeDate = t
			END IF
		ELSE
			encodeDate = s
		END IF
	END IF
	'encodeDate = CSTR( CDATE( s ) )
END FUNCTION












]]></msxsl:script>

</xsl:stylesheet>