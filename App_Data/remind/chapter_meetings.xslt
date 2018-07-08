<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions"
>

<xsl:output method="xml" version="1.0" indent="yes" encoding="ISO-8859-1" 
cdata-section-elements="subject content location comments" />


<xsl:template match="/|@*|node()">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="category[.='c-meeting']">
<category>meeting</category>
</xsl:template>

<xsl:template match="category[.='email2']">
<category>neighbor-meeting</category>
</xsl:template>

<xsl:template match="@time" />

<xsl:template match="event[category='al-e']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>


<xsl:template match="event[category='al-h']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="category[.='al-h']">
<category>email</category>
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="style[parent::event/category='al-h']">
<style>RmdChapterNeighbor</style>
</xsl:template>

<xsl:template match="event[category='al-y']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="style[parent::event/category='al-y']">
<style>RmdChapterNeighborSpecial</style>
</xsl:template>


<xsl:template match="category[.='al-y']">
<category>email</category>
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="event[category='al-k']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>


<xsl:template match="event[category='al-n']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>

<xsl:template match="event[category='al-s']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>


<xsl:template match="event[category='al-y']">
<xsl:copy>
<xsl:apply-templates select="@*|node()"  />
</xsl:copy>
</xsl:template>


<xsl:template match="event" />




</xsl:stylesheet>