<!--#include file="tab_tools.asp"-->
<%


CLASS CTabFormatGray

	PROPERTY GET colorBackground()
		colorBackground = "#999999"
	END PROPERTY
	
	PROPERTY GET colorTab()
		colorTab = "#CCCCCC"
	END PROPERTY
	
	PROPERTY GET colorTabSelected()
		colorTabSelected = "#FFFFFF"
	END PROPERTY
	
	PROPERTY GET classTab()
		classTab = "shoppingtab"
	END PROPERTY
	
	PROPERTY GET classSelected()
		classSelected = ""
	END PROPERTY
	
	PROPERTY GET alignTabHoriz()
		alignTabHoriz = "center"
	END PROPERTY
	
	PROPERTY GET imageTL()
		imageTL = "images/pie_tl_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageTR()
		imageTR = "images/pie_tr_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageBL()
		imageBL = "images/pie_bl_gray.gif"
	END PROPERTY

END CLASS






CLASS CTabShoppingData

	PRIVATE m_oRS
	PRIVATE m_aData
	PRIVATE m_i
	PRIVATE m_sURL
	PRIVATE	m_oTabID
	PRIVATE m_oName
		

	PRIVATE SUB Class_Initialize
		SET m_oTabID = Nothing
		SET m_oName = Nothing
		SET m_oRS = Nothing
		m_sURL = ""
	END SUB
	
	PRIVATE SUB Class_Terminate
		SET m_oTabID = Nothing
		SET m_oName = Nothing
		SET m_oRS = Nothing
	END SUB

	
	PROPERTY GET RecordCount()
		RecordCount = m_oRS.RecordCount
	END PROPERTY
	
	PROPERTY GET EOF()
		EOF = m_oRS.EOF
	END PROPERTY
	
	SUB MoveFirst()
		m_oRS.MoveFirst
	END SUB
	
	SUB MoveNext()
		m_oRS.MoveNext
	END SUB
	
	FUNCTION IsCurrent( x )
		IF LCASE(x) = LCASE(recString( m_oTabID )) THEN
			IsCurrent = TRUE
		ELSE
			IsCurrent = FALSE
		END IF
	END FUNCTION
	
	PROPERTY GET HREF()
		HREF = "list.asp?tab=" & m_oTabID
	END PROPERTY
	
	PROPERTY GET TabLabel()
		TabLabel = recString( m_oName )
	END PROPERTY
	
	'----------------
	
	PROPERTY LET URL( s )
		m_sURL = s
	END PROPERTY
	
	PROPERTY SET RS( o )
		SET m_oRS = o
		SET m_oName = m_oRS.Fields("Name")
		SET m_oTabID = m_oRS.Fields("TabID")
	END PROPERTY
	
END CLASS



SUB outputShoppingTabs

	DIM	oRS

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"TabID, " _
		&		"Name " _
		&	"FROM " _
		&		"Tabs " _
		&	"WHERE " _
		&		"0 = Disabled " _
		&		"AND " _
		&			"0 < (SELECT Count(*) " _
		&			"FROM Products " _
		&				"INNER JOIN TabCategoryMap ON Products.Category = TabCategoryMap.CategoryID " _
		&			"WHERE TabCategoryMap.TabID = Tabs.TabID AND 0=Products.Disabled) " _
		&	"ORDER BY " _
		&		"IF( Tabs.SortName IS NULL,'',Tabs.SortName)+Tabs.Name " _
		&	";"
	
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF 0 < oRS.RecordCount THEN
		DIM	oTabGen
		DIM	oTabData
		DIM	oTabFormat
		
		SET oTabGen = NEW CTabGenerate
		SET oTabData = NEW CTabShoppingData
		SET oTabFormat = NEW CTabFormatGray
		
		SET oTabData.RS = oRS
		oTabData.URL = "tab_tools_test.asp?tab="
		
		SET oTabGen.TabData = oTabData
		SET oTabGen.TabFormat = oTabFormat
		oTabGen.Current = g_sTab
		oTabGen.makeTabs
		
		
		SET oTabFormat = Nothing
		SET oTabData = Nothing
		SET oTabGen = Nothing


	END IF
	'IF 6 < oRS.RecordCount THEN
	'	outputShoppingTabsVert oRS
	'ELSEIF 0 < oRS.RecordCount THEN
	'	outputShoppingTabsHoriz oRS
	'END IF
	oRS.Close
	SET oRS = Nothing


END SUB

SUB outputShoppingTabsHoriz( oRS )
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" bgcolor="#999999" height="6"><spacer type="block" height="1" width="1"></td>
  </tr>
  <tr>
    <td width="100%" align="center" bgcolor="#999999">
<%

	IF 0 < oRS.RecordCount THEN

		DIM	oTabID
		DIM	oName
		
		SET oTabID = oRS.Fields("TabID")
		SET oName = oRS.Fields("Name")
		
		DIM	sBGColor
		DIM	bThisID
		
		Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCRLF
		Response.Write "<tr>" & vbCRLF
		Response.Write "<td width=""3""></td>" & vbCRLF
		DO UNTIL oRS.EOF
			IF CSTR(g_sTab) = recString( oTabID ) THEN
				bThisID = TRUE
				sBGColor = "#FFFFFF"
			ELSE
				bThisID = FALSE
				sBGColor = "#CCCCCC"
			END IF
			Response.Write "<td valign=""top"" bgcolor=""" & sBGColor & """>"
			Response.Write "<img src=""images/pie_tl_gray.gif"" width=""8"" height=""8"">"
			Response.Write "</td>" & vbCRLF
			
			Response.Write "<th bgcolor=""" & sBGColor & """>"
			IF bThisID THEN
				Response.Write Server.HTMLEncode(recString( oName ))
			ELSE
				Response.Write "<a class=""shoppingtab"" href=""list.asp?tab=" & oTabID
				Response.Write """>"
				Response.Write Server.HTMLEncode(recString( oName ))
				Response.Write "</a>"
			END IF
			Response.Write "</th>" & vbCRLF
			
			Response.Write "<td valign=""top"" align=""right"" bgcolor=""" & sBGColor & """>"
			Response.Write "<img src=""images/pie_tr_gray.gif"" width=""8"" height=""8"">"
			Response.Write "</td>" & vbCRLF
			
			Response.Write "<td width=""3""></td>" & vbCRLF

			oRS.MoveNext
		LOOP
		
		SET oTabID = Nothing
		SET oName = Nothing
		Response.Write "</tr>" & vbCRLF
		Response.Write "</table>" & vbCRLF
	END IF
%>
    </td>
  </tr>
</table>
<%
END SUB


SUB outputShoppingTabsVert( oRS )
%>
<table border="0" cellspacing="0" cellpadding="0" align="left">
  <tr>
    <td bgcolor="#999999" width="6"><spacer type="block" height="1" width="1"></td>
    <td bgcolor="#999999" align="right">
<%

	IF 0 < oRS.RecordCount THEN
	
		DIM	oTabID
		DIM	oName
		
		SET oTabID = oRS.Fields("TabID")
		SET oName = oRS.Fields("Name")
		
		DIM	sBGColor
		DIM	bThisID
		
		Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCRLF
		Response.Write "<tr>" & vbCRLF
		Response.Write "<td height=""8"" colspan=""2""></td>" & vbCRLF
		Response.Write "</tr>" & vbCRLF
		DO UNTIL oRS.EOF
			Response.Write "<tr>" & vbCRLF
			IF CSTR(g_sTab) = recString( oTabID ) THEN
				bThisID = TRUE
				sBGColor = "#FFFFFF"
			ELSE
				bThisID = FALSE
				sBGColor = "#CCCCCC"
			END IF
			Response.Write "<td valign=""top"" bgcolor=""" & sBGColor & """>"
			Response.Write "<img src=""images/pie_tl_gray.gif"" width=""8"" height=""8"">"
			Response.Write "</td>" & vbCRLF
			
			Response.Write "<th bgcolor=""" & sBGColor & """ align=""right"" rowspan=""2"">"
			IF bThisID THEN
				Response.Write Server.HTMLEncode(recString( oName ))
			ELSE
				Response.Write "<a class=""shoppingtab"" href=""list.asp?tab=" & oTabID
				Response.Write """>"
				Response.Write Server.HTMLEncode(recString( oName ))
				Response.Write "</a>"
			END IF
			Response.Write "</th>" & vbCRLF
			Response.Write "<td bgcolor=""" & sBGColor & """ rowspan=""2"">"
			Response.Write "&nbsp;"
			Response.Write "</td>" & vbCRLF
			Response.Write "</tr>" & vbCRLF
			
			Response.Write "<tr>" & vbCRLF
			Response.Write "<td valign=""bottom"" align=""left"" bgcolor=""" & sBGColor & """>"
			Response.Write "<img src=""images/pie_bl_gray.gif"" width=""8"" height=""8"">"
			Response.Write "</td>" & vbCRLF
			Response.Write "</tr>" & vbCRLF
			
			Response.Write "<tr>" & vbCRLF
			Response.Write "<td height=""3"" colspan=""2""></td>" & vbCRLF
			Response.Write "</tr>" & vbCRLF

			oRS.MoveNext
		LOOP
		
		SET oTabID = Nothing
		SET oName = Nothing
		Response.Write "</tr>" & vbCRLF
		Response.Write "</table>" & vbCRLF
	END IF


%>
    </td>
    <td width="8" valign="top" align="left"><img src="images/pie_tl_gray.gif" width="8" height="8"></td>
  </tr>
  <tr>
    <td bgcolor="#999999" width="6"><spacer type="block" height="1" width="1"></td>
    <td bgcolor="#999999" align="right" valign="bottom"><img src="images/pie_br.gif" width="8" height="8"></td>
    <td width="8" valign="top" align="left"></td>
</table>
<%
END SUB




SUB outputPad
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" height="6"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<%
END SUB




'outputShoppingTabs
'outputPad

%>