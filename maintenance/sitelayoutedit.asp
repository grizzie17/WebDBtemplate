<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )

%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\index_tools.asp"-->
<%


DIM	g_sType
	g_sType = Request("type")

	IF "" = g_sType THEN g_sType = "screen"

	DIM	g_sContentTitle
	DIM	g_sLayout
	DIM	g_sConfigName
	DIM	g_sOutName

	IF "mobile" = g_sType THEN
		g_sContentTitle = "Mobile Homepage Content"
		g_sLayout = "dynamic"
		g_sConfigName = "config_mobile.txt"
		g_sOutName = "default_body_mobile.asp"
	ELSE
		g_sContentTitle = "Homepage Content"
		g_sLayout = "fixed1,dynamic,fixed2"
		g_sConfigName = "config.txt"
		g_sOutName = "default_body.asp"
	END IF

SUB makeBlocks()
	DIM	aSplit
	DIM	sLine
	FOR EACH sLine IN aFileList
		aSplit = SPLIT( sLine, vbTAB )
		Response.Write "<div class=""block"" id=""" & aSplit(kFI_Name) & """>" & vbCRLF
		Response.Write "<h1 class=""draghandle"">" & Server.HTMLEncode(aSplit(kFI_Title)) & "</h1>" & vbCRLF
		Response.Write "<p>" & Server.HTMLEncode(aSplit(kFI_Description)) & "</p>" & vbCRLF
		Response.Write "</div>"
	NEXT
END SUB

DIM	g_aCol(4)

SUB makeData()


	DIM	sFolderLayout
	sFolderLayout = findAppDataFolder( "layout\homepage" )

	DIM	sFolderConfig
	sFolderConfig = g_oFSO.BuildPath( sFolderLayout, "config" )

	DIM	i
	FOR i = 0 TO UBOUND(g_aCol)
		g_aCol(i) = ""
	NEXT

	DIM	sData
	DIM	aSplit
	DIM	sLine
	DIM	sFileConfig
	sFileConfig = sFolderConfig & "\" & g_sConfigName
	IF g_oFSO.FileExists( sFileConfig ) THEN
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sFileConfig )
		IF NOT oFile IS Nothing THEN
			i = 1
			DO UNTIL oFile.AtEndOfStream
				sLine = oFile.ReadLine
				aSplit = SPLIT( sLine, ":" )
				IF "layout" = aSplit(0) THEN
					g_sLayout = aSplit(1)
				ELSE
					sData = RIGHT( aSplit(0), 1 )
					IF ISNUMERIC( sData ) THEN
						g_aCol(CLNG(sData)) = aSplit(1)
					ELSE
						g_aCol(i) = aSplit(1)
						i = i + 1
					END IF
				END IF
			LOOP
			oFile.Close
			SET oFile = Nothing
		END IF
	END IF

	DIM	sUsed
	sUsed = "," & JOIN( g_aCol, "," ) & ","
	sData = ""
	FOR EACH sLine IN aFileList
		aSplit = SPLIT( sLine, vbTAB )
		IF INSTR( ","&aSplit(kFI_NAME)&",", sUsed ) < 1 THEN
			sData = sData & ", " & "'" & aSplit(kFI_Name) & "'"
		END IF
	NEXT
	DIM	aData(4)
	DIM	j
	FOR i = LBOUND(g_aCol) TO UBOUND(g_aCol)
		sLine = g_aCol(i)
		aSplit = SPLIT(sLine,",")
		FOR j = LBOUND(aSplit) TO UBOUND(aSplit)
			aSplit(j) = "'" & aSplit(j) & "'"
		NEXT
		aData(i) = JOIN(aSplit,",")
	NEXT
	sData = MID(sData, 2)	' trim leading comma
	Response.Write "'column-0': [" & sData & "]," & vbCRLF
	Response.Write "'column-1': [" & aData(1) & "]," & vbCRLF
	Response.Write "'column-2': [" & aData(2) & "]," & vbCRLF
	Response.Write "'column-3': [" & aData(3) & "]" & vbCRLF
END SUB

SUB makeLayoutSelect()
	DIM	i
	DIM	aSplit
	FOR i = LBOUND(aLayoutList) TO UBOUND(aLayoutList)
		aSplit = SPLIT( aLayoutList(i), vbTAB )
		Response.Write "<option "
		IF aSplit(kFI_Description) = g_sLayout THEN
			Response.Write "selected=""selected"" "
		END IF
		Response.Write "value=""" & aSplit(kFI_Description) & """>" & Server.HTMLEncode(aSplit(kFI_Title)) & "</option>" & vbCRLF
	NEXT
END SUB


g_sUseFileNameSuffix = ".txt"

DIM	sFolderLayout
sFolderLayout = findAppDataFolder( "layout\homepage\layouts" )
IF "" <> sFolderLayout THEN
	buildFileList virtualFromPhysicalPath(sFolderLayout), FALSE

	DIM	aLayoutList()
	REDIM aLayoutList(UBOUND(aFileList))

	DIM	ix
	FOR ix = LBOUND(aFileList) TO UBOUND(aFileList)
		aLayoutList(ix) = aFileList(ix)
	NEXT
END IF

DIM	sFolderComponents
sFolderComponents = findAppDataFolder( "layout\homepage\components" )
IF "" <> sFolderComponents THEN
	buildFileList virtualFromPhysicalPath(sFolderComponents), FALSE
END IF



%>
<!DOCTYPE html>
<html>
	<head>
		<title>Edit Layouts</title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		
		<script type="text/javascript" src="js/scriptaculous-js/lib/prototype.js"></script>
		<script type="text/javascript" src="js/scriptaculous-js/src/scriptaculous.js"></script>
		<script type="text/javascript" src="js/portal.js"></script>
		
		<!--? require_once 'blocks_get.php'; ?-->
		
		<script type="text/javascript">
			var settings = {
				<% makeData %>
			};
			
			var options 	= { portal 			: 'columns', 
				editorEnabled 	: true /*, 
				saveurl 		: 'blocks_save.php'*/ };
								
			var data 		= {};
						
			var portal;
					
			Event.observe(window, 'load', function() {
				portal = new Portal(settings, options, data);
			});
		</script>
		<script type="text/javascript">

			function doCancel()
			{
				window.location.href = "site.asp";
			}


			function updateColumns()
			{
				var x = document.getElementById("data-layout");
				if ( x )
				{
					var	s = x.value;
					var a = s.split(",");
					var n = a.size();
					var oColumn1 = document.getElementById("column-1");
					var oColumn2 = document.getElementById("column-2");
					var oColumn3 = document.getElementById("column-3");
					if ( 1 < n )
					{
						oColumn1.className = "column " + a[0];
						oColumn1.style.display = "block";
						oColumn2.className = "column " + a[1];
						oColumn2.style.display = "block";
						if ( 2 < n )
						{
							oColumn3.className = "column " + a[2];
							oColumn3.style.display = "block";
						}
						else
						{
							oColumn3.style.display = "none";
						}
					}
					else
					{
						oColumn1.className = "column dynamic";
						oColumn1.style.display = "block";
						oColumn2.style.display = "none";
						oColumn3.style.display = "none";
					}
				}
			}

			function changeLayout(o)
			{
				var	idx = o.selectedIndex;
				var	s = o.options[idx].value;
				var	x = document.getElementById("data-layout");
				if ( x )
					x.value = s;
				updateColumns();
				return false;
			}

		</script>
		
		<link rel="stylesheet" href="css/style.css"  type="text/css" media="screen" />
		<link rel="stylesheet" href="css/portal.css"  type="text/css" media="screen" />
	</head>
	<body>
		<script type="text/javascript">
			window.onload = function () {
				updateColumns();
			};
		</script>
	
		<div id="wrapper">
			
			<div id="heading">
				<div id="title-0">Widgets</div>
				<div id="title-1"><%=g_sContentTitle %></div>
			</div>
			
			<!-- This is the part you want -->
			
			<div id="columns">
				
				<div id="column-0" class="column store"></div>

				<div id="column-1" class="column fixed1"></div>
				
				<div id="column-2" class="column dynamic"></div>
				
				<div id="column-3" class="column fixed2"></div>
				
				
				<div class="portal-column" id="portal-column-block-list" style="display: none;">
					<%
						makeBlocks
					%>
				</div>
				
			</div>
			
			<!-- Above is the part you want -->
			
						

			<div class="box" style="clear:both;">
				<form action="sitelayouteditsubmit.asp" method="post" id="layoutform">
					<select name="layout-select" onchange="changeLayout(this)">
						<%makeLayoutSelect %>

					</select>&nbsp;&nbsp;&nbsp;
					<input type="hidden" name="configfile" value="<%=g_sConfigName %>" />
					<input type="hidden" name="outfile" value="<%=g_sOutName %>" />
					<input type="hidden" name="layout" id="data-layout" value="<%=g_sLayout %>" />
					<input type="hidden" name="column-0" id="data-column-0" value="" />
					<input type="hidden" name="column-1" id="data-column-1" value="<%=g_aCol(1) %>" />
					<input type="hidden" name="column-2" id="data-column-2" value="<%=g_aCol(2) %>" />
					<input type="hidden" name="column-3" id="data-column-3" value="<%=g_aCol(3) %>" />
					<input type="submit" name="B1" value="Save Layout" />&nbsp;&nbsp;&nbsp;
					<input type="button" value="Cancel" name="B2" onclick="doCancel()" />
				</form>
			</div>
			
			
			<div style="clear:both;"></div>	
			
			
		</div>
		
	</body>
</html>
<%
	SET g_oFSO = Nothing
%>