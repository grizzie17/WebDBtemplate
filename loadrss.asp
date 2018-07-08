<%@ Language=VBScript %>
<%
OPTION EXPLICIT

DIM	g_oFSO
SET	g_oFSO = Server.CreateObject("Scripting.FileSystemObject")
%>
<!--#include file="scripts/findfiles.asp"-->
<!--#include file="scripts/rss.asp"-->
<!--#include file="scripts/include_cache.asp"-->
<%



DIM	q_sFile
DIM	q_sURL
DIM	q_xFail

q_sFile = Request("f")
q_sURL = Request("h")
q_xFail = Request("x")


SUB AMADirectLink

	ON ERROR Goto 0
	
	
	DIM	oXML
	SET oXML = rss_getXMLCached( q_sFile, q_sURL, "", "h", 4, "d" )
	IF NOT Nothing IS oXML THEN
	
		IF NOT Nothing IS oXML.selectSingleNode("//feed") THEN
			DIM	otXML
			DIM	otXSL
			SET otXSL = Server.CreateObject("msXML2.DOMDocument.6.0")
			otXSL.async = FALSE
			otXSL.resolveExternals = false
			otXSL.setProperty "ProhibitDTD", false
			otXSL.setProperty "ResolveExternals", true
			otXSL.setProperty "AllowDocumentFunction", true
			otXSL.setProperty "AllowXsltScript", true
			otXSL.load( Server.MapPath( "Scripts/atom2rss.xml" ) )
			IF otXSL.parseError.errorCode <> 0 THEN
				Response.Write "Error in XSLT file:<br>" & vbCRLF
				Response.Write "Error Code: " & otXSL.parseError.errorCode & "<br>" & vbCRLF
				Response.Write "Error Reason: " & otXSL.parseError.reason & "<br>" & vbCRLF
				Response.Write "Error Line: " & otXSL.parseError.line & "<br>" & vbCRLF
			END IF
			otXML = oXML.transformNode(otXSL)
			oXML.loadxml( otXML )
		END IF

		DIM oXSL
		SET oXSL = Server.CreateObject("msxml2.DOMDocument.6.0")
		oXSL.async = FALSE
		oXSL.resolveExternals = false
		oXSL.setProperty "ProhibitDTD", false
		oXSL.setProperty "ResolveExternals", true
		oXSL.setProperty "AllowDocumentFunction", true
		oXSL.setProperty "AllowXsltScript", true
		oXSL.load( Server.MapPath( "scripts/rssama.xslt" ) )
		IF oXSL.parseError.errorCode <> 0 THEN
			Response.Write "Error in XSLT file:<br>" & vbCRLF
			Response.Write "Error Code: " & oXSL.parseError.errorCode & "<br>" & vbCRLF
			Response.Write "Error Reason: " & oXSL.parseError.reason & "<br>" & vbCRLF
			Response.Write "Error Line: " & oXSL.parseError.line & "<br>" & vbCRLF
		END IF

		DIM sText
		sText = oXML.transformNode(oXSL)
		IF 0 < LEN(sText) THEN
			Response.Write "<div class=""rssama"">" & vbCRLF
			Response.Write sText
			Response.Write "</div>" & vbCRLF
		END IF
		
		SET oXSL = Nothing
		SET oXML = Nothing
	
	ELSE
		IF "" <> q_xFail THEN
			Response.Write "<a target=""_blank"" href=""http://" & q_xFail & """>" & Server.HTMLEncode(q_xFail) & "</a>"
		ELSE
			Response.Write q_sFile
		END IF

	END IF
	
END SUB
AMADirectLink






SET g_oFSO = Nothing













%>