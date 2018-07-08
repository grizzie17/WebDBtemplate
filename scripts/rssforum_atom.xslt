<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet
	version="1.0"
	exclude-result-prefixes="msxsl local"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:local="#local-functions">

<xsl:output method="html" indent="no" encoding="ISO-8859-1"/>

<xsl:template match="/rss">
		<xsl:apply-templates select="channel"/>
</xsl:template>

<xsl:template match="channel">
	<xsl:choose>
		<xsl:when test="item">
				<xsl:apply-templates select="item"/>
		</xsl:when>
	</xsl:choose>
</xsl:template>

<xsl:template match="pubDate">
	<br/><span class="rssForumPubDate">
	Updated: <xsl:value-of select="local:encodeDate(.)"/>
	</span>
</xsl:template>

<xsl:template match="lastBuildDate">
	<div style="color: #999999; font-size: smaller; font-family: sans-serif;">
	Date: <xsl:value-of select="local:encodeDate(.)"/>
	</div>
</xsl:template>

<xsl:template match="description">
	<span class="rssForumDescription">
	<xsl:value-of disable-output-escaping="yes" select="."/>
	</span>
</xsl:template>

<xsl:template match="image">
	<a><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute>
	<img border="0" alt="." align="absbottom"><xsl:attribute name="src"><xsl:value-of select="url"/></xsl:attribute></img></a>
</xsl:template>

<xsl:template match="item">
	<xsl:if test="local:testDateRange(pubDate)">
			<li class="rssForumTopic">
			<a target="_blank"><xsl:attribute name="href"><xsl:value-of select="link"/></xsl:attribute><xsl:value-of disable-output-escaping="yes" select="local:topictitle(title)"/></a><br/>
			_$$nbsp;_$$nbsp; <span style="color: #999999; font-size: smaller; font-family: sans-serif"><xsl:value-of select="local:DateRange(pubDate)"/> in <xsl:value-of select="category"/></span>
			</li>
	</xsl:if>
</xsl:template>


<xsl:template match="enclosure">
	<br/><span class="rssEnclosure">
	<img ><xsl:attribute name="src"><xsl:value-of select="@url"/></xsl:attribute></img>
	</span>
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


FUNCTION testDateRange( oNode )
	testDateRange = false
	DIM	s
	s = oNode.item(0).text
	DIM	a
	a = SPLIT(s,",")
	s = TRIM(a(1))
	a = SPLIT(s," ")
	s = a(0) & "-" & a(1) & "-" & a(2)
	IF ISDATE( s ) THEN
		DIM	d
		d = CDATE( s )
		IF DATEDIFF( "d", d, NOW ) < 6 THEN
			testDateRange = true
		END IF
	END IF
END FUNCTION

FUNCTION DateRange( oNode )
	DateRange = ""
	DIM	s
	s = oNode.item(0).text
	DIM	a
	a = SPLIT(s,",")
	s = TRIM(a(1))
	a = SPLIT(s," ")
	s = a(0) & "-" & a(1) & "-" & a(2)
	IF ISDATE( s ) THEN
		DIM	d
		d = CDATE( s )
		DIM	n
		n = DATEDIFF( "d", d, NOW )
		IF n < 1 THEN
			DateRange = "today"
		ELSEIF 1 = n THEN
			DateRange = "yesterday"
		ELSE
			DateRange = n & " days ago"
		END IF
	END IF
END FUNCTION

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
					t = t & " " & a(i)
				NEXT 'i
				IF ISDATE( t ) THEN
					d = CDATE( t )
					encodeDate = formatPostDate( d )
				ELSE
					t = a(1) & " " & a(0) & ", " & a(2)
					FOR i = 3 TO UBOUND(a)
						t = t & " " & a(i)
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