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
<!--#include file="scripts/rssweather.asp"-->
<%


FUNCTION fetchHTTP( sURL )
	fetchHTTP = ""
	DIM	oHTTP
	SET oHTTP = xmldom_makeHTTP()
	IF NOT oHTTP IS Nothing THEN
		oHTTP.Open "GET", sURL, False
		oHTTP.SetRequestHeader "User-Agent", Request.ServerVariables("HTTP_USER_AGENT")
		oHTTP.Send
		IF 0 = Err.Number THEN
			DIM	nStatus
			nStatus = CINT(oHTTP.Status)
			IF 200 = nStatus  OR  0 = nStatus THEN
			
				fetchHTTP = oHTTP.ResponseText
			ELSE
				fetchHTTP = "Error " & nStatus & "<br>" & vbCRLF
				fetchHTTP = fetchHTTP & Server.HTMLEncode( sURL ) & vbCRLF
			END IF
		END IF
	END IF
	SET oHTTP = Nothing
END FUNCTION



FUNCTION remoteExtras( sName, sHeader )
	DIM	sURL
	DIM	sDOM
	IF "" <> g_sDistrictDomain THEN
		sDOM = g_sDistrictDomain
	ELSE
		sDOM = Request.ServerVariables("HTTP_HOST")
	END IF
	sURL = "http://" & sDOM & "/srv_extras.asp?q=" & sName & "&h=" & Server.URLEncode(sHeader)
	remoteExtras = fetchHTTP( sURL )
END FUNCTION



FUNCTION getExtras( sName, nTime, sHeader )

	DIM	sData
	sData = ""
	DIM	sCacheFile
	sCacheFile = LCASE(sName) & ".htm"
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, "n", nTime, "d" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sData = remoteExtras( sName, sHeader )
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", sCacheFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
	getExtras = sData
	
END FUNCTION



DIM	q_sQ
DIM	q_sT
DIM q_sH
DIM	nTime
nTime = 30

q_sQ = Request("q")
IF "" = q_sQ THEN
	q_sQ = "Humor"
END IF
q_sT = Request("t")
IF "" <> q_sT THEN
	IF ISNUMERIC(q_sT) THEN
		nTime = CLNG(q_sT)
	END IF
END IF
q_sH = Request("h")
IF "" <> q_sQ THEN
	DIM	sData
	sData = getExtras( q_sQ, nTime, q_sH )
	IF 0 < LEN(sData) THEN
		Response.Status = "200"
		Response.ContentType = "text/html"
		Response.Write sData
	ELSE
		Response.Status = "404 Not Found"
        Response.Write "404 Not Found" & vbCRLF
	END IF
ELSE
	Response.Status = "403 Forbidden"
	Response.Write "403 Forbidden" & vbCRLF
END IF

Response.Flush
Response.End



SET g_oFSO = Nothing




%>