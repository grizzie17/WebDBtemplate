<%


FUNCTION rssannounce_makexml()
	SET rssannounce_makexml = xmldom_make()
END FUNCTION


FUNCTION rssannounce_filter()

	DIM	s
	s = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCRLF
	s = s & "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & vbCRLF
	s = s & "<xsl:output method=""xml"" version=""1.0"" indent=""yes"" encoding=""ISO-8859-1"" />" & vbCRLF
	s = s & vbCRLF
	s = s & "<xsl:template match=""/|@*|node()"">" & vbCRLF
	s = s & "<xsl:copy>" & vbCRLF
	s = s & "<xsl:apply-templates select=""@*|node()""  />" & vbCRLF
	s = s & "</xsl:copy>" & vbCRLF
	s = s & "</xsl:template>" & vbCRLF
	s = s & vbCRLF
	
	DIM	sCat
	DIM	aCat
	
	IF "" <> g_sRSSAnnounceInclude THEN
		sCat = REPLACE(g_sRSSAnnounceInclude, " ", "" )
		aCat = SPLIT(sCat, ",")
		FOR EACH sCat IN aCat
			IF "none" = sCat THEN
				s = s & "<xsl:template match=""item[category]"">" & vbCRLF
			ELSE
				s = s & "<xsl:template match=""item[category='" & sCat & "']"">" & vbCRLF
			END IF
			s = s & "<xsl:copy>" & vbCRLF
			s = s & "<xsl:apply-templates select=""@*|node()""  />" & vbCRLF
			s = s & "</xsl:copy>" & vbCRLF
			s = s & "</xsl:template>" & vbCRLF
			s = s & vbCRLF
		NEXT
	END IF
	
	IF "" <> g_sRSSAnnounceExclude THEN
		sCat = REPLACE(g_sRSSAnnounceExclude, " ", "" )
		aCat = SPLIT(sCat, ",")
		FOR EACH sCat IN aCat
			IF "none" = sCat THEN
				s = s & "<xsl:template match=""item[category]"" />" & vbCRLF
			ELSE
				s = s & "<xsl:template match=""item[category='" & sCat & "']"" />" & vbCRLF
			END IF
			s = s & vbCRLF
		NEXT
	END IF
	
	IF "" <> g_sRSSAnnounceInclude THEN
		s = s & "<xsl:template match=""item"" />" & vbCRLF
	END IF
	s = s & vbCRLF
	s = s & "</xsl:stylesheet>" & vbCRLF
	
	rssannounce_filter = s

END FUNCTION



FUNCTION rssannounce_loadFilter()

	DIM	oXSLT
	SET oXSLT = rssannounce_makexml()

	DIM	sCacheFile
	sCacheFile = cache_checkFile( "rss", "rssannouncefilter.xslt", "m", 6, "y" )
	IF "" = sCacheFile THEN
		DIM	sFilterXml
		sFilterXml = rssannounce_filter()
		
		DIM	oFile
		SET oFile = cache_makeFile( "rss", "rssannouncefilter.xslt" )
		IF NOT oFile IS Nothing THEN
			oFile.Write sFilterXml
			oFile.Close
			SET oFile = Nothing
		END IF
		oXSLT.loadxml( sFilterXml )
	ELSE
		oXSLT.Load sCacheFile
	END IF
	SET rssannounce_loadFilter = oXSLT
	SET oXSLT = Nothing

END FUNCTION


FUNCTION rssannounce_fetch( sLocalFile, sRemoteURL )

	DIM	oXML
	SET oXML = Nothing
	DIM	sCacheFile
	sCacheFile = cache_checkFile( "rss", sLocalFile, "h", 12, "d" )
	IF "" = sCacheFile THEN
		DIM	sContent
		DIM	sMime
		DIM	bBinary
		
		rss_fetchHTTP sContent, sMime, sRemoteURL, "", bBinary
		IF 20 < LEN(sContent) THEN

			DIM	oXFilter
			SET oXFilter = rssannounce_loadFilter()
			
			DIM	oTXML
			SET oTXML = rssannounce_makexml()
			oTXML.loadxml(sContent)
			
			SET oXML = rssannounce_makexml()
			DIM	sResultXML
			sResultXML = oTXML.transformNode(oXFilter)
			DIM oFile
			SET oFile = cache_makeFile( "rss", sLocalFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sResultXML
				oFile.Close
				SET oFile = Nothing
			END IF
			
			oXML.loadxml(sResultXML)
			
			SET oTXML = Nothing
			SET oXFilter = Nothing
			
		END IF
	ELSE
		SET oXML = rssannounce_makexml()
		oXML.load sCacheFile
	END IF
	SET rssannounce_fetch = oXML
	SET oXML = Nothing



END FUNCTION



SUB RSSDistrictAnnouncements

	DIM	oXML
	SET oXML = rssannounce_fetch( "districtannounce.xml", "http://" & g_sDistrictDomain & "/rssannouncements.asp" )
	IF NOT oXML IS Nothing THEN
	
		DIM	oXSL
		SET oXSL = xmldom_loadFile( Server.MapPath( "scripts/rssannounce.xslt" ))
		IF NOT oXSL IS Nothing THEN
			DIM sText
			sText = oXML.transformNode(oXSL)
			IF 0 < LEN(sText) THEN
				sText = REPLACE( sText, "_$$", "&" )
				Response.Write sText
			END IF
			SET oXSL = Nothing
		END IF
		
		SET oXML = Nothing
	ELSE
		Response.Write "Unable to fetch district announcements<br>"

	END IF
	
END SUB






%>