<?xml version="1.0"?>
<!--DOCTYPE xsl:stylesheet
[
	<!ENTITY nbsp	"&#160;" >
	<!ENTITY raquo	"&amp;raquo;" >
]-->
 
<xsl:stylesheet
	version="1.0"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions"
	exclude-result-prefixes="msxsl local"
	>

<xsl:output method="html" indent="no" encoding="ISO-8859-1"/>


<xsl:template match="/rss">
		<xsl:apply-templates select="channel"/>
</xsl:template>

<xsl:template match="channel">
	<xsl:choose>
		<xsl:when test="item">
			<div class="rssForumTitle">
			<xsl:choose>
				<xsl:when test="image">
			<xsl:apply-templates select="image"/><br/>
				</xsl:when>
			</xsl:choose>
			<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute><xsl:value-of disable-output-escaping="yes" select="title"/></a>
			</div>
			<!--xsl:apply-templates select="description"/-->
			<xsl:choose>
				<xsl:when test="pubDate">
			<xsl:apply-templates select="pubDate"/>
				</xsl:when>
				<xsl:otherwise>
			<xsl:apply-templates select="lastBuildDate"/>
				</xsl:otherwise>
			</xsl:choose>
			<table cellpadding="0" cellspacing="0" border="0">
		
			<xsl:apply-templates select="item"/>
			
			</table>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<xsl:template match="pubDate">
	<xsl:if test="local:uniqueDate(.)">
	<div style="color: #999999; font-size: smaller; font-family: sans-serif;">
	Updated: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
	</xsl:if>
</xsl:template>

<xsl:template match="lastBuildDate">
	<div style="color: #999999; font-size: smaller; font-family: sans-serif;">
	Date: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
</xsl:template>

<xsl:template match="description">
	<div>
	<xsl:value-of disable-output-escaping="yes" select="."/>
	</div>
</xsl:template>

<xsl:template match="image">
	<a><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute>
	<img border="0" alt="." align="absbottom"><xsl:attribute name="src"><xsl:value-of select="url"/></xsl:attribute></img></a>
</xsl:template>

<xsl:template match="item">
	<tr>
	<xsl:choose>
		<xsl:when test="local:istopic(title)">
			<td valign="top" nowrap="nowrap">_$$nbsp;_$$nbsp;_$$raquo;_$$nbsp;_$$nbsp;</td>
			<td>
			<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute><xsl:value-of disable-output-escaping="yes" select="local:topictitle(title)"/></a>
			<xsl:if test="title != description  and  link != description">
			<div class="rssItemDescription">
			<xsl:value-of disable-output-escaping="yes" select="description"/>
			</div>
			</xsl:if>
			<xsl:apply-templates select="enclosure"/>
			<xsl:apply-templates select="pubDate"/>
			</td>
		</xsl:when>
		<xsl:otherwise>
			<th align="left" colspan="2">
			<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute><xsl:value-of disable-output-escaping="yes" select="local:topictitle(title)"/></a>
			</th>
		</xsl:otherwise>
	</xsl:choose>
	</tr>
</xsl:template>


<xsl:template match="enclosure">
	<div class="rssEnclosure">
	<img ><xsl:attribute name="src"><xsl:value-of select="@url"/></xsl:attribute></img>
	</div>
</xsl:template>


<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[



function istopic(oNode)

	DIM	s
	s = oNode.item(0).text
	IF "-- " = LEFT( s, 3 ) THEN
		istopic = TRUE
	ELSE
		istopic = FALSE
	END IF

end function


function topictitle(oNode)

	DIM	s
	s = oNode.item(0).text
	IF "-- " = LEFT( s, 3 ) THEN
		topictitle = MID(s,4)
	ELSE
		topictitle = s
	END IF

end function


function encodeText(oNode)
	DIM s
	s = oNode.item(0).text
	''s = REPLACE( s, "<", "&lt;", 1, -1, vbTextCompare )
	''s = REPLACE( s, ">", "&gt;", 1, -1, vbTextCompare )
	''s = REPLACE( s, "&", "&amp;", 1, -1, vbTextCompare )
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
	'EXIT FUNCTION
	IF ISDATE( s ) THEN
		d = CDATE( s )
		encodeDate = FORMATDATETIME( d, 2 ) & " " & FORMATDATETIME( d, 4 )
	ELSE
		DIM	i
		i = INSTR( s, "," )
		IF 0 < i THEN
			DIM	t
			t = TRIM(MID( s, i+1 ))
			IF ISDATE( t ) THEN
				d = CDATE( t )
				encodeDate = CSTR(FORMATDATETIME( d, 2 ))
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