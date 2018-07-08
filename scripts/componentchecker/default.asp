<%@Language="VBScript"%>
<%
OPTION EXPLICIT



response.buffer=true


DIM	aCategoryMap
aCategoryMap = ARRAY( _
		"Miscellaneous", _
		"Email", _
		"Browser", _
		"Upload", _
		"Image", _
		"Documents", _
		"FileManagement", _
		"Charts", _
		"Server", _
		"Users", _
		"E-Commerce", _
		"Validation", _
		"Forms", _
		"XML" _
		)


%>
<html>
<head>
<TITLE>componentchecker.asp</TITLE>
<meta name="robots" content="noindex">
<base target="_blank">
</head>
<body bgcolor="#FFFFFF">
<%

SUB getComponentInfo( bInstalled, sIcon, sTitle, sExtra, sComponent, sLine )
' format: comObject|comURL|comName|comCategory|comCategory2

	bInstalled = FALSE
	sIcon = "unknown"
	sTitle = ""
	sExtra = ""
	sComponent = ""
	IF "" <> sLine THEN
		DIM	nTemp
		DIM	aSplit
		aSplit = SPLIT( sLine, "|" )
		IF 0 < UBOUND(aSplit) THEN
			sComponent = aSplit(0)
			sTitle = aSplit(2)
			nTemp = aSplit(3)
			IF ISNUMERIC(nTemp) THEN
				sIcon = aCategoryMap(nTemp)
			ELSEIF "" <> nTemp THEN
				sIcon = nTemp
			END IF
			IF "" = sTitle THEN sTitle = sComponent
		ELSE
			sComponent = sLine
			sTitle = sComponent
		END IF
		IF "" <> sComponent THEN
			DIM	o
			DIM	nErr
			ON ERROR Resume Next
			SET o = Server.CreateObject( sComponent )
			nErr = Err.Number
			ON ERROR Goto 0
			IF 0 = nErr THEN
				bInstalled = TRUE
				
				DIM	sVersion
				sVersion = ""
				sExpires = ""
				
				ON ERROR Resume Next
				sVersion = TRIM(o.version)
				IF "" = sVersion THEN sVersion = TRIM(o.getversion)
				IF "" = sVersion THEN sVersion = TRIM(o.getvers)
				IF "" = sVersion THEN sVersion = TRIM(o.vers)
				sExpires = TRIM(o.expires)
				ON ERROR Goto 0
				
				IF "unavailable" = LCASE(sVersion) THEN sVersion = ""
				IF "unavailable" = LCASE(sExpires) THEN sExpires = ""
				IF "unknown" = LCASE(sExpires) THEN sExpires = ""
				IF "n/a" = LCASE(sExpires) THEN sExpires = ""
				IF "9/9/9999" = LCASE(sExpires) THEN sExpires = "Never"
				IF "" <> sVersion THEN
					IF "" <> sExpires THEN
						sExtra = "(version: " & sVersion & "; expires: " & sExpires & ")"
					ELSE
						sExtra = "(version: " & sVersion & ")"
					END IF
				ELSEIF "" <> sExpires THEN
					sExtra = "(expires: " & sExpires & ")"
				END IF
			ELSE
				bInstalled = FALSE
				IF -2147221005 <> nErr THEN
					sExtra = "<font color=""red"">failed. Error #" & nErr & "</font>"
				END IF
			END IF
			SET o = Nothing
		END IF
	END IF
END SUB





SUB beginComponentGroup( sLine )
	DIM	sData
	IF "[" = LEFT(sLine,1) THEN
		DIM	sGroup
		DIM	sGroupLC
		DIM	sTemp
		
		sGroup = MID( sLine, 2, LEN(sLine)-2 )
		sGroupLC = LCASE( sGroup )
		IF 1 = INSTR(sGroupLC,"http:") THEN
			sData = "<a href=""" & sGroup & """>" & Server.HTMLEncode(sGroup) & "</a>"
		ELSEIF 1 = INSTR(sGroupLC, "www.") THEN
			sData = "<a href=""http://" & sGroup & """>" & Server.HTMLEncode(sGroup) & "</a>"
		ELSE
			sData = sGroup
		END IF
	ELSE
		sData = sLine
	END IF
%>
<table cellpadding="2" style="border-collapse: collapse" width="500" border="1" bordercolor="#C0C0C0">
	<tr><th align="left" bgcolor="#C0C0C0"><%=sData%></td></tr>
	<tr>
		<td>
	<!--blockquote-->
<%
END SUB




SUB addComponent( sIcon, sComponent, sTitle, sExtra, bInstalled )
%>
<img src="<%=sIcon%>.gif"> &nbsp;
<%
	IF bInstalled THEN
%>
	<b><%=sComponent%></b> <%=sExtra%><br>
<%
	ELSE
%>
	<font color="#CCCCCC"><%=sComponent%></font> <%=sExtra%><br>
<%
	END IF
END SUB



SUB endComponentGroup()
%>
	<!--/blockquote-->
	</td>
</tr>
</table>
<br>
<%
	Response.Flush
END SUB



dim successSTR, FailSTR, checkSTR


DIM	whichfile
DIM	fs
DIM	thisfile

whichfile=server.mappath("component.ini")
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.OpenTextFile(whichfile, 1, False)

DIM	sExpires

DIM	sExtra
DIM	bInstalled
DIM	sTitle
DIM	sComponent
DIM	sIcon



beginComponentGroup "IIS Server Information"
addComponent "bullet", ScriptEngine & " Version: ", "", ScriptEngineMajorVersion() & "." & ScriptEngineMinorVersion, TRUE
addComponent "bullet", "Server Name: ", "", Request.ServerVariables("SERVER_NAME"), TRUE
addComponent "bullet", "Server Software: ", "", Request.ServerVariables("SERVER_SOFTWARE"), TRUE


DIM	thisline
DIM	attempt

DIM counter
DIM	successcount
DIM	failcount

counter=0
successcount = 0
failcount = 0


DO UNTIL thisfile.AtEndOfStream
	counter=counter+1
	thisline=thisfile.readline
	attempt=trim(thisline)
	DO WHILE attempt="" AND thisfile.atendofStream=false
		thisline=thisfile.readline
		attempt=trim(thisline)
	LOOP
	
	IF "[" = LEFT(attempt,1) THEN

		endComponentGroup
		beginComponentGroup attempt

	ELSEIF "" <> attempt THEN
		IF "!" <> LEFT(attempt,1) THEN
'SUB getComponentInfo( bInstalled, sIcon, sTitle, sExtra, sComponent, sLine )
			getComponentInfo bInstalled, sIcon, sTitle, sExtra, sComponent, attempt
			addComponent sIcon, sComponent, sTitle, sExtra, bInstalled
			If bInstalled then
				
				successcount=successcount+1
			else
				failcount=failcount+1
			end if
		END IF
	END IF
loop
endComponentGroup

thisfile.Close
set thisfile=nothing
set fs=nothing

'response.write "Checked: " & checkSTR & "<hr>"
'response.write successSTR & "<hr>"
'response.write failSTR 
%>

</body>

</html>