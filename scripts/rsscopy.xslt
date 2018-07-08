<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
version="1.0">

<xsl:output method="xml" version="1.0" indent="yes" encoding="ISO-8859-1" />

<xsl:template match="/ | @* | node()">
<xsl:copy>
<xsl:apply-templates select="@* | node()"  />
</xsl:copy>
</xsl:template>


</xsl:stylesheet>