<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_xmldom.asp"-->
<!--#include file="scripts\remind_utils.asp"-->
<%


DIM	q_q
DIM	q_h	'href
DIM	q_i	'interval
DIM	q_v	'value
DIM	q_b	'break-interval

q_q = Request("q")
q_h = Request("h")
q_i = Request("i")
q_v = Request("v")
q_b = Request("b")

IF "" = q_i THEN q_i = "d"
IF "" = q_v THEN q_v = 5
IF "" = q_b THEN
	SELECT CASE q_i
	CASE "h"
		IF q_v < 24 THEN
			q_b = "d"
		ELSEIF q_v < 24*7 THEN
			q_b = "w"
		ELSEIF q_v < 24*7*4 THEN
			q_b = "m"
		ELSE
			q_b = "y"
		END IF
	CASE "d"
		IF q_v < 7 THEN
			q_b = "w"
		ELSEIF q_v < 28 THEN
			q_b = "m"
		ELSE
			q_b = "m"
		END IF
	CASE "w"
		IF q_v < 4 THEN
			q_b = "m"
		ELSE
			q_b = "y"
		END IF
	CASE "m"
		q_b = "y"
	CASE ELSE
		q_b = "y"
	END SELECT
END IF



FUNCTION getHTTP( sHREF )
	getHTTP = ""
	DIM oHTTP
	SET oHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	IF NOT oHTTP IS Nothing THEN
		'Response.Write sHREF & "<br>"
		'Response.Flush
		oHTTP.Open "GET", sHREF, False
		oHTTP.SetRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")
		oHTTP.Send
		IF 0 = Err.Number THEN
			DIM	nStatus
			nStatus = CINT(oHTTP.Status)
			IF 200 = nStatus  OR  0 = nStatus THEN
			
				'Response.Write "nStatus = " & nStatus & "<br>"
				'Response.Flush
			
				DIM	sNewMime
				
				sNewMime = oHTTP.getResponseHeader( "Content-Type" )
				'Response.Write "Mime = " & sNewMime & "<br>"
				SELECT CASE LCASE(sNewMime)
				CASE "text/xml"
					getHTTP = oHTTP.ResponseXML.xml
				END SELECT
							
			END IF
		END IF
	
		SET oHTTP = Nothing
	END IF
END FUNCTION


DIM	sXMLData
sXMLData = ""

IF "" <> q_q  AND  "" <> q_h THEN
	DIM	sFile
	sFile = q_q & ".xml"

	DIM	oFile
	SET oFile = cache_openTextFile( "remind", sFile, q_i, q_v, q_b )
	IF NOT oFile IS Nothing THEN
		sXMLData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		DIM	sHREF
		IF INSTR( q_h, "://" ) < 1 THEN
			q_h = "http://" & q_h
		END IF
		IF "/" <> RIGHT( q_h, 1 ) THEN
			q_h = q_h & "/"
		END IF
		sHREF = q_h & "remindserver.asp?q=" & q_q
		sXMLData = getHTTP( sHREF )
		IF "" <> sXMLData THEN
			DIM	sXsltFile
			sXsltFile = findRemindFile( q_q & ".xslt" )
			IF "" <> sXsltFile THEN
				DIM	oXML
				DIM	oXSLT
				SET oXML = xmlDom_loadXML( sXMLData )
				IF NOT oXML IS Nothing THEN
					SET oXSLT = xmlDom_LoadFile( sXsltFile )
					IF NOT oXSLT IS Nothing Then
						sXMLData = oXML.transformNode(oXSLT)
					END IF
				END IF
				SET oXML = Nothing
				SET oXSLT = Nothing
			END IF
			SET oFile = cache_makeFile( "remind", sFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sXMLData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
END IF

IF "" <> sXMLData THEN
	Response.ContentType = "text/xml"
	Response.Write sXMLData
	
ELSE
	Response.Status = "404 File Not Found"
	Response.Write "File Not Found"
END IF
Response.Flush
Response.End



SET g_oFSO = Nothing
%>