<%@ Language=VBScript %>
<%
OPTION EXPLICIT

DIM	g_oFSO
SET	g_oFSO = Server.CreateObject("Scripting.FileSystemObject")
%>
<!--#include file="scripts/config.asp"-->
<!--#include file="config/configuser.asp"-->
<!--#include file="scripts/findfiles.asp"-->
<!--#include file="scripts/rss.asp"-->
<!--#include file="scripts/include_xmldom.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<!--#include file="scripts/include_rssannounce.asp"-->
<%


FUNCTION genAnnouncements()

	genAnnouncements = ""
	DIM	oXML
	SET oXML = rssannounce_fetch( "districtannouncements.xml", "http://" & g_sDistrictDomain & "/rssannouncements.asp" )
	IF NOT oXML IS Nothing THEN
		DIM	oXSL
		SET oXSL = xmldom_loadFile( Server.MapPath( "scripts/rssannounce.xslt" ))
		IF NOT oXSL IS Nothing THEN
			DIM sText
			sText = oXML.transformNode(oXSL)
			IF 0 < LEN(sText) THEN
				sText = REPLACE( sText, "_$$", "&" )
				genAnnouncements = sText
			END IF
			SET oXSL = Nothing
		END IF
		
		SET oXML = Nothing
	END IF
	
END FUNCTION



FUNCTION getAnnouncements()

	DIM	sData
	sData = ""
	DIM	sCacheFile
	sCacheFile = "districtannouncements.htm"
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, "h", 6, "d" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sData = genAnnouncements()
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", sCacheFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		ELSE
			getAnnouncements = "Unable to fetch district announcements<br>"
		END IF
	END IF
	getAnnouncements = sData
	
END FUNCTION


DIM	sData
sData = getAnnouncements()
IF 0 < LEN(sData) THEN
	Response.ContentType = "text/text"
	Response.Write sData
ELSE
	Response.Status = "404 Not Found"
END IF

Response.Flush
Response.End




SET g_oFSO = Nothing




%>