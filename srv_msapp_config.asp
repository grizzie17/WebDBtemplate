<%@	Language=VBScript%>
<%
OPTION EXPLICIT


	DIM	g_oFSO
	SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")
%>
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%

FUNCTION findImage( sPath, sName )
	findImage = ""

	DIM sFile
	sFile = g_oFSO.BuildPath( sPath, sName & ".png" )
	IF g_oFSO.FileExists( sFile ) THEN
		findImage = virtualFromPhysicalPath( sFile )
		EXIT FUNCTION
	END IF

	sFile = g_oFSO.BuildPath( sPath, sName & ".gif" )
	IF g_oFSO.FileExists( sFile ) THEN
		findImage = virtualFromPhysicalPath( sFile )
		EXIT FUNCTION
	END IF

END FUNCTION


FUNCTION buildConfig()
	DIM	sXML
	sXML = ""

	DIM	oFolder
	DIM	sPath
	
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))
	sPath = oFolder.Path
	SET oFolder = Nothing

	DIM	s70
	DIM	s150
	DIM	s310x150
	DIM	s310

	s70 = findImage( sPath, "favicon128" )
	s150 = findImage( sPath, "favicon270" )
	s310x150 = findImage( sPath, "favicon588x270" )
	s310 = findImage( sPath, "favicon588" )

	IF 0 < (LEN(s70) + LEN(s150) + LEN(s310x150) + LEN(s310)) THEN

		sXML = sXML & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCRLF
		sXML = sXML & "<browserconfig>" & vbCRLF
		sXML = sXML & "<msapplication>" & vbCRLF
		sXML = sXML & "<tile>" & vbCRLF
		IF 0 < LEN(s70) THEN
			sXML = sXML & "<square70x70logo src=""" & s70 & """/>" & vbCRLF
		END IF
		IF 0 < LEN(s150) THEN
			sXML = sXML & "<square150x150logo src=""" & s150 & """/>" & vbCRLF
		END IF
		IF 0 < LEN(s310x150) THEN
			sXML = sXML & "<square310x150logo src=""" & s310x150 & """/>" & vbCRLF
		END IF
		IF 0 < LEN(s310) THEN
			sXML = sXML & "<square310x310logo src=""" & s310 & """/>" & vbCRLF
		END IF
		sXML = sXML & "</tile>" & vbCRLF

		IF 0 < (LEN(s150) + LEN(s310x150) + LEN(s310)) THEN
			DIM	sHost
			sHost = Request.ServerVariables("HTTP_HOST")
	
			DIM s
			DIM i
			s = Request.ServerVariables("URL")
			i = INSTRREV( s, "/" )
			IF 0 < i THEN
				s = LEFT( s, i-1 )
			ELSE
				s = ""
			END IF
			DIM	sLivetile
			sLivetile = "http://" & sHost & s & "/srv_msapp_livetile.asp"
			sXML = sXML & "<notification>" & vbCRLF
			sXML = sXML & "<polling-uri src=""" & sLivetile & "?q=1"" />" & vbCRLF
			sXML = sXML & "<polling-uri2 src=""" & sLivetile & "?q=2"" />" & vbCRLF
			sXML = sXML & "<polling-uri3 src=""" & sLivetile & "?q=3"" />" & vbCRLF
			sXML = sXML & "<frequency>360</frequency>" & vbCRLF
			sXML = sXML & "<cycle>1</cycle>" & vbCRLF
			sXML = sXML & "</notification>" & vbCRLF
		END IF

		sXML = sXML & "</msapplication>" & vbCRLF
		sXML = sXML & "</browserconfig>" & vbCRLF

	END IF

	buildConfig = sXML
END FUNCTION

FUNCTION getConfig()
	CONST kFilename = "browserconfig.xml"
	DIM	sXML
	DIM	oFile
	SET oFile = cache_openTextFile( "site", kFilename, "h", 8, "d" )
	IF NOT oFile IS Nothing THEN
		sXML = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sXML = buildConfig()
		IF 0 < LEN(sXML) THEN
			SET oFile = cache_makeFile( "site", kFilename )
			IF NOT oFile IS Nothing THEN
				oFile.Write sXML
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
	getConfig = sXML
END FUNCTION

	DIM sXML
	'sXML = getConfig()
	sXML = buildConfig()
	IF 0 < LEN(sXML) THEN

		Response.ContentType = "text/xml"
		Response.Write sXML

	END IF


%>
<% 
	SET g_oFSO = Nothing
%>