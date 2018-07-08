<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions"
	>

<xsl:output method="html" indent="no" encoding="ISO-8859-1"/>


<xsl:template match="/rss">
		<xsl:apply-templates select="channel"/>
</xsl:template>

<xsl:template match="channel">
	<xsl:apply-templates select="item"/>
</xsl:template>

<xsl:template match="item">
			<xsl:apply-templates select="title" />
			<div class="BlockBody">
			<xsl:apply-templates select="enclosure" />
			<xsl:apply-templates select="description" />
			<xsl:apply-templates select="pubDate" />
			</div>
</xsl:template>

<xsl:template match="pubDate">
	<div class="rssAnnouncementsPubDate">
	Updated: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
</xsl:template>

<xsl:template match="lastBuildDate">
	<div class="rssAnnouncementsLastBuildDate">
	Date: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
</xsl:template>

<xsl:template match="title">
	<div class="BlockHead">
	<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="../link" /></xsl:attribute>
	<xsl:value-of disable-output-escaping="yes" select="."/></a>
	</div>
</xsl:template>

<xsl:template match="description">
	<div class="rssAnnouncementsDescription">
	<xsl:value-of disable-output-escaping="yes" select="."/>
	_$$nbsp;_$$nbsp; 
	<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="../link" /></xsl:attribute>more...</a>
	</div>
</xsl:template>


<xsl:template match="image">
	<xsl:choose>
		<xsl:when test="link">
		<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute>
		<xsl:call-template name="image_content"/>
		</a>
		</xsl:when>
		<xsl:otherwise>
		<xsl:call-template name="image_content"/>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template name="image_content">
	<img border="0" align="absbottom">
	<xsl:if test="title">
		<xsl:attribute name="alt"><xsl:value-of select="title"/></xsl:attribute>
	</xsl:if>
	<xsl:attribute name="src"><xsl:value-of select="url"/></xsl:attribute></img>
</xsl:template>




<xsl:template match="enclosure">
	<div class="rssAnnouncementsEnclosure">
	<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="../link" /></xsl:attribute>
	<img border="0" ><xsl:attribute name="src"><xsl:value-of select="@url"/></xsl:attribute></img></a>
	</div>
</xsl:template>




<msxsl:script implements-prefix="local" language="VBscript">
<![CDATA[


DIM	g_nest
g_nest = 0


function nest()
	g_nest = g_nest + 1
	nest = ""
end function

function unnest()
	IF 0 < g_nest THEN
		g_nest = g_nest - 1
	END IF
	unnest = ""
end function

function isnest()
	IF 0 < g_nest THEN
		isnest = TRUE
	ELSE
		isnest = FALSE
	END IF
end function

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
		encodeDate = formatPostDate( d )
	ELSE
		DIM	i
		i = INSTR( s, "," )
		IF 0 < i THEN
			DIM	t
			t = TRIM(MID( s, i+1 ))
			IF ISDATE( t ) THEN
				d = CDATE( t )
				encodeDate = formatPostDate( d )
			ELSE
				DIM	a
				a = SPLIT( t, " " )
				'REDIM PRESERVE a(UBOUND(a)-1)
				t = a(0) & "-" & a(1) & "-" & a(2)
				FOR i = 3 TO UBOUND(a)
					IF "UTC" <> a(i) THEN
						t = t & " " & a(i)
					END IF
				NEXT 'i
				IF ISDATE( t ) THEN
					d = CDATE( t )
					encodeDate = formatPostDate( d )
				ELSE
					t = a(1) & " " & a(0) & ", " & a(2)
					FOR i = 3 TO UBOUND(a)
						IF "UTC" <> a(i) THEN
							t = t & " " & a(i)
						END IF
					NEXT 'i
					IF ISDATE( t ) THEN
						d = CDATE( t )
						encodeDate = formatPostDate( d )
					ELSE
						encodeDate = t
					END IF
				END IF
			END IF
		ELSE
			encodeDate = s
		END IF
	END IF
	'encodeDate = CSTR( CDATE( s ) )
END FUNCTION



FUNCTION formatPostDate( d )

	DIM	sD
	
	sD = MONTH(d) & "/" & RIGHT("0"&DAY(d),2) & "/" & YEAR(d)
	
	DIM	sSfx
	sSfx = "am"
	DIM	nH
	DIM	nM
	DIM	nS
	nH = HOUR( d )
	nM = MINUTE( d )
	nS = SECOND( d )
	'nH = CINT( MID(sDate,9,2) )
	IF nH < 1 THEN
		nH = 12
		IF 0 = nM  AND  0 = nS THEN sSfx = "midnight"
	ELSEIF 11 < nH THEN
		sSfx = "pm"
		IF 0 = nM  AND  0 = nS THEN sSfx = "noon"
		IF 12 < nH THEN nH = nH - 12
	END IF
	
	formatPostDate = sD & " " & RIGHT("0"&nH,2) & ":" & RIGHT("0"&nM,2) & ":" & RIGHT("0"&nS,2) & " " & sSfx

END FUNCTION










]]></msxsl:script>

</xsl:stylesheet>