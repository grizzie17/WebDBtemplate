<%
'---------------------------------------------------------------------
'            Copyright 1986 .. 2007 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'---------------------------------------------------------------------

OPTION EXPLICIT

Response.Expires= 0

%>
<script language="JavaScript" runat="SERVER">
function URLDecode( s )
{
	return unescape( s );
}
</script>
<%

DIM	oFSO
DIM	sTemp
DIM	sQueryRaw
DIM	sQuery
DIM	sURLRaw
DIM	sURL
DIM	sFile
DIM	i
	


SUB beginHTML
	DIM	iTemp
	DIM	sHost
	DIM	sRootDir
	
	sTemp = Request.ServerVariables("URL")
	iTemp = INSTRREV( sTemp, "/", -1, vbTextCompare )
	IF 1 < iTemp THEN
		sRootDir = LEFT(sTemp,iTemp)
	ELSE
		sRootDir = "/"
	END IF

	sHost = "http://" & Request.ServerVariables("HTTP_HOST") & sRootDir
%>
<html>

<head>
<title>404 Not Found - RocketCityWings.org</title>
<base href="<%=sHost%>">
<meta name="robots" content="noindex">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body         { background-color: #FFFFFF }
b, strong    { font-weight: bold }
i, em        { font-style: italic }
p            { margin-left: 8 }
.navigationpath { color: #000000; background-color: #FFCC33; font-size: 10px; font-weight: bold; 
               font-family: Arial, Helvetica, Verdana, sans-serif }
.titleblock  { color: #000000; background-color: #FFFFFF; background-image: 
               url('images/bkg_title.jpg'); font-weight: bold; font-family: 
               Arial, Helvetica, sans-serif; padding-left: 10px }
.navbackground { background-color: #663300; color: #FFFFFF }
-->
</style>
</head>

<body topmargin="0" leftmargin="0" marginheight="0" marginwidth="0"
bgcolor="#FFFFFF">

<%
END SUB

SUB outputHeader( sMessage )
	IF FALSE THEN
%>
<p>=====Begin Header=====</p>
<%
	END IF
%>

<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" class="titleblock">
      <h2>&nbsp;<%=sMessage%></h2>
    </td>
    <td align="right" valign="bottom">
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="bottom">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
              <tr>
                <td width="100%" bgcolor="#993300"><img border="0"
                  src="images/pie_brx.gif"></td>
              </tr>
            </table>
          </td>
          <td bgcolor="#993300">
			<img border="0"
            src="images/rcw_small.gif" width="100" height="86"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#993300" height="3"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="2">
  <tr>
    <th bgcolor="#CC9966" align="right" class="navigationpath"><span
      class="navigationpath"><a class="navigationpath" href="./">Rocket City 
	Wings (Home)</a></span></th>
  </tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#993300" height="1"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<%
	IF FALSE THEN
%>
<p>=====End Header=====</p>
<%
	END IF
	IF Response.Buffer THEN Response.Flush
END SUB

SUB makeHeader
	outputHeader "Requested file was Not Found"
END SUB

SUB makeSiteDiabled
	outputHeader "Site Temporarily Disabled"
END SUB



SUB makeFooter
	IF FALSE THEN
%>
<p>=====Begin Footer=====</p>
<%
	END IF
%>
<hr>
<p>You may also want to look at the following links:</p>
<ul>
  <li><a href="./">RocketCityWings.org</a> - Home</li>
</ul>

<%
	IF FALSE THEN
%>
<p>=====End Footer=====</p>
<%
	END IF
	IF Response.Buffer THEN Response.Flush
END SUB 'makeFooter


SUB patchWebLinks
	DIM	sFileLC
	DIM	sQueryLC
	DIM	sTemp
	DIM	i,j
	
	sFileLC = LCASE(sFile)
	IF 0 < LEN(sQuery) THEN
		sQuery = URLDecode( sQuery )
		i = INSTR(1,sQuery,"&",vbTextCompare)
		IF 0 < i THEN sQuery = LEFT(sQuery,i-1)
		i = INSTR( 1, sQuery, "://", vbTextCompare )
		IF 0 < i THEN
			i = INSTR( i+3, sQuery, "/", vbTextCompare )
			IF 0 < i THEN
				sQuery = MID(sQuery,i)
			'ELSE
			'	sQuery = ""
			END IF
		ELSE
			IF "/" <> LEFT(sQuery,1) THEN sQuery = ""
		END IF
		i = INSTRREV( sFileLC, "/", -1, vbTextCompare )
		sFileLC = LEFT( sFileLC, i-1 )
		sFileLC = sFileLC & LCASE(sQuery)
		'Response.Write "sFileLC= " & sFileLC & "<br>" & vbCRLF
		sQuery = ""
	END IF
	
	i = INSTR(1,sFileLC,"regional/",vbTextCompare)
	IF 0 < i THEN
		'Response.Write "Regional/ - encountered<br>" & vbCRLF
		j = INSTRREV(sFileLC,"/",-1,vbTextCompare)
		sTemp = MID(sFileLC,j+1)
		SELECT CASE sTemp
		CASE "1intro.htm"
			sTemp = "index.asp"
		CASE "educate.htm"
			sTemp = "education.asp"
		CASE "tv.htm"
			sTemp = "television.htm"
		CASE "weather_stormwatch.htm"
			sTemp = "weather_warnings.asp"
		END SELECT
		sQueryLC = MID(sFileLC,i,j-i+1)
		sFileLC = LEFT(sFileLC,i-1) & sTemp
		SELECT CASE sQueryLC
		CASE "regional/al/hsv/"
			sQuery = "loc=al_huntsville"
		CASE "regional/ga/atl/"
			sQuery = "loc=ga_atlanta"
		CASE "regional/ga/savannah/"
			sQuery = "loc=ga_savannah"
		CASE "regional/mo/stlouis"
			sQuery = "loc=mo_stlouis"
		CASE "regional/nc/chrlt/"
			sQuery = "loc=nc_charlotte"
		CASE "regional/sc/gvl/"
			sQuery = "loc=sc_greenville"
		CASE "regional/tn/nash/"
			sQuery = "loc=tn_nashville"
		CASE "regional/tx/hou/"
			sQuery = "loc=tx_houston"
		CASE ELSE
			sQuery = ""
		END SELECT
		'Response.Write "sQuery = " & sQuery & "<br>" & vbCRLF
	END IF
	
	sFile = REPLACE(sFileLC,"weblinks/","links/",1,-1,vbTextCompare)
	sFile = REPLACE(sFile,"/links/links/","/links/", 1, -1, vbTextCompare )
	sFile = REPLACE(sFile,"~grizzie/","", 1, -1, vbTextCompare )
	IF "/" = MID(sFile,LEN(sFile)) THEN sFile = sFile & "index.asp"
	'Response.Write "sFile = " & sFile & "<br>" & vbCRLF
	
END SUB


SUB patchFile
    DIM	sFileLC
	sFileLC = LCASE(sFile)
	IF 0 < INSTR(1,sFileLC,"/weblinks/",vbTextCompare) THEN patchWebLinks
END SUB

SUB patchIMG
	DIM sFileLC
	sFileLC = LCASE(sFile)
	IF 0 = INSTR(1,sFileLC,"/links/",vbTextCompare) THEN
		IF "/" = LEFT(sFile,1) THEN
			sFile = "/links" & sFile
		ELSE
			sFile = "/links/" & sFile
		END IF
	END IF
END SUB


'===========================================================================



SUB makeRedirect( sHREF )
%>
<p><a href="<%=sHREF%>">Loading File</a></p>
<%
END SUB 'makeRedirect


SUB makeNotFound
%>

<p><b>The requested file:</b><br>
<%=Server.HTMLEncode(sURLRaw)%></p>
<%
	
	IF Response.Buffer THEN Response.Flush
	IF 0 < LEN(sFile)  AND  "/" <> RIGHT(sFile,1) THEN
		sPossible = Server.MapPath( sFile )
		IF NOT oFSO.FileExists( sPossible ) THEN
			patchFile
			i = INSTRREV( sFile, ".", -1, vbTextCompare )
			IF 0 < i THEN
				sFile = LEFT(sFile,i-1)
				sPossible = Server.MapPath( sFile )
			END IF
			IF oFSO.FileExists( sPossible & ".asp" ) THEN
				sFile = sServer & sFile & ".asp"
			ELSEIF oFSO.FileExists( sPossible & ".htm" ) THEN
				sFile = sServer & sFile & ".htm"
			ELSEIF oFSO.FileExists( sPossible & ".html" ) THEN
				sFile = sServer & sFile & ".html"
			ELSEIF oFSO.FileExists( sPossible & ".shtm" ) THEN
				sFile = sServer & sFile & ".shtm"
			ELSEIF oFSO.FileExists( sPossible & ".shtml" ) THEN
				sFile = sServer & sFile & ".shtml"
			ELSE
				sFile = ""
			END IF
		END IF
		IF 0 < LEN(sFile) THEN
			IF 0 < LEN(sQuery) THEN
				sHREF = sFile & "?" & sQuery
				sFile = sFile & "?" & sQuery
			ELSE
				sHREF = sFile
			END IF
      %>
<p><font size="4"><b>A possible match may be:</b><br>
<a href="<%=sHREF%>"><%=Server.HTMLEncode(sFile)%></a></font></p>
      <%
		END IF
	END IF
	IF Response.Buffer THEN Response.Flush

END SUB


SUB makeQueryRequired
%>
<p>This file must be used from server error processing mechanisms.</p>
<%
END SUB


SUB endHTML
%>

</body>

</html>
<%
	IF Response.Buffer THEN Response.Flush
END SUB


PUBLIC aImgSplit
FUNCTION findIMG( sFile )
	DIM j
	DIM i
	DIM k
	DIM sPath
	findIMG = ""
	IF "/" = LEFT(sFile,1) THEN
		sPath = MID(sFile,2)
		aImgSplit = SPLIT( sPath, "/", -1, vbTextCompare )
	ELSE
		aImgSplit = SPLIT( sFile, "/", -1, vbTextCompare )
	END IF
	IF 1 < UBOUND(aImgSplit) THEN
		FOR j = UBOUND(aImgSplit) TO 1 STEP-1
			FOR i = j-1 TO 0 STEP-1
				sPath = "/"
				FOR k = 0 TO i-1
					sPath = sPath & aImgSplit(k) & "/"
				NEXT 'k
				FOR k = j to UBOUND(aImgSplit)-1
					sPath = sPath & aImgSplit(k) & "/"
				NEXT 'k
				sPath = sPath & aImgSplit(UBOUND(aImgSplit))
				IF oFSO.FileExists( Server.MapPath( sPath )) THEN
					findImg = sPath
					EXIT FUNCTION
				END IF
			NEXT 'i
		NEXT 'j
	END IF
	IF oFSO.FileExists( Server.MapPath( sFile )) THEN findImg = sFile
END FUNCTION


SUB listSV()
%>
<table border="2" cellspacing="1" cellpadding="2">
<%
    DIM sName
    
    FOR EACH sName IN Request.ServerVariables
%>
  <tr>
    <th align="left"><%=sName%></th>
    <td><%=Server.HTMLEncode(Request.ServerVariables(sName))%></td>
  </tr>
<%
    NEXT 'sName
%>
</table>
<%
END SUB
'listSV

sQueryRaw = Request.ServerVariables("QUERY_STRING")
IF 0 = LEN(sQueryRaw) THEN
	beginHTML
	makeQueryRequired
	endHTML
ELSE
	i = INSTR(1,sQueryRaw,";",vbTextCompare)
	IF 0 < i THEN
		sURLRaw = MID(sQueryRaw,i+1)
	ELSE
		sURLRaw = sQueryRaw
	END IF
	IF INSTR( 1, sURLRaw, "%", vbTextCompare ) THEN sURLRaw = URLDecode( sURLRaw )

	i = INSTR( 1, sURLRaw, "?", vbTextCompare )
	IF 0 < i THEN
		sFile = LEFT( sURLRaw, i-1 )
		sQuery = MID( sURLRaw, i+1 )
	ELSE
		i = INSTR( 1, sURLRaw, "&", vbTextCompare )
		IF 0 < i THEN
			sFile = LEFT( sURLRaw, i-1 )
			sQuery = MID( sURLRaw, i+1 )
		ELSE
			sFile = sURLRaw
			sQuery = ""
		END IF
	END IF
	i = INSTR( 1, sFile, "|", vbTextCompare )
	IF 0 < i THEN
		sFile = LEFT( sFile, i-1 )
	END IF
	IF 0 < LEN(sQuery) THEN
		sTemp = sQuery
		sTemp = REPLACE( sTemp, "<b>", "", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "</b>", "", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "%3Cb%3E", "", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "%3C/b%3E", "", 1, -1, vbTextCompare )
		sTemp = Server.URLEncode(sTemp)
		sTemp = REPLACE( sTemp, "%26", "&", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "%3D", "=", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "%3d", "=", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "%5F", "_", 1, -1, vbTextCompare )
		sTemp = REPLACE( sTemp, "%5f", "_", 1, -1, vbTextCompare )
		sQuery = sTemp
	END IF

	' separate the server and filename
	DIM	sServer
	i = INSTR( 1, sFile, "://", vbTextCompare )
	IF 0 < i THEN
		i = INSTR( i+3, sFile, "/", vbTextCompare )
		IF 0 < i THEN
			sServer = LEFT(sFile,i-1)
			sFile = MID(sFile,i)
		ELSE
			sServer = sFile
			sFile = ""
		END IF
	ELSE
		sServer = ""
		IF "/" <> LEFT(sFile,1) THEN sFile = ""
	END IF

	sTemp = REPLACE( sFile, "'", "", 1, -1, vbTextCompare )
	sTemp = REPLACE( sTemp, "<b>", "", 1, -1, vbTextCompare )
	sTemp = REPLACE( sTemp, "</b>", "", 1, -1, vbTextCompare )
	sFile = REPLACE( sTemp, """", "", 1, -1, vbTextCompare )
	sTemp = ""
	IF 0 < LEN(sFile) THEN
		i = INSTRREV(sFile,"/",-1,vbTextCompare)
		IF 0 < i THEN
			sTemp = MID(sFile,i+1)
			i = INSTRREV(sTemp,".",-1,vbTextCompare)
			IF 0 < i THEN
				sTemp = LCASE(MID(sTemp,i))
			END IF
		END IF
	END IF


	DIM	sHREF
	DIM	sPossible
	
	SET oFSO = Server.CreateObject("Scripting.FileSystemObject")
	IF 0 < LEN(sTemp) THEN
		IF INSTR( 1, sURLRaw, "news:", vbTextCompare ) THEN
			i = INSTR( 1, sURLRaw, "news:", vbTextCompare )
			IF 0 < i THEN
				sHREF = MID( sURLRaw, i )
				Response.Redirect sHREF
				
				beginHTML
				makeRedirect sHREF
				endHTML
			END IF
		ELSEIF INSTR( 1, ".asp.htm.html.stm.shtm.shtml", sTemp, vbTextCompare ) THEN
		
			sPossible = Server.MapPath( sFile )
			IF oFSO.FileExists( sPossible ) THEN
				IF 0 < LEN(sQuery) THEN
					sHREF = sFile & "?" & sQuery
				ELSE
					sHREF = sFile
				END IF
				
				Response.Redirect sHREF
				
				beginHTML
				makeRedirect sHREF
				endHTML
			ELSE
				Response.Status = "404 Not Found"

				beginHTML
				makeHeader
				makeNotFound
				makeFooter
				endHTML
			END IF
			
		ELSEIF INSTR( 1, ".gif.jpg.jpeg.png.css.js", sTemp, vbTextCompare ) THEN
			' special lookup if the file is an image

			patchIMG
			sTemp = findIMG( sFile )
			IF 0 = LEN(sTemp) THEN
				Response.Status = "404 Not Found"
				Response.Write "404 Not Found"
			ELSE
				Response.Redirect sTemp
			END IF
		ELSEIF INSTR( 1, ".ico", sTemp, vbTextCompare ) THEN
			Response.Status = "404 Not Found"
			Response.Write "404 Not Found"
		ELSE
			Response.Status = "404 Not Found"

			beginHTML
			makeHeader
			makeNotFound
			makeFooter
			endHTML
		END IF
	ELSE
		sPossible = Server.MapPath( sFile )
		IF oFSO.FolderExists( sPossible ) THEN
			IF 0 < LEN(sQuery) THEN
				sHREF = sFile & "?" & sQuery
			ELSE
				sHREF = sFile
			END IF

			Response.Redirect sHREF

			beginHTML
			makeRedirect sHREF
			endHTML
		ELSE
			Response.Status = "404 Not Found"
			beginHTML
			makeHeader
			makeNotFound
			makeFooter
			endHTML
		END IF
	END IF
END IF


SET oFSO = Nothing
Response.End
%>