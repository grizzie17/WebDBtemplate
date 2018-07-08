<%


CLASS CTabFormatDummy

	PROPERTY GET colorBackground()
		colorBackground = "#999999"
	END PROPERTY
	
	PROPERTY GET colorTab()
		colorTab = "#CCCCCC"
	END PROPERTY
	
	PROPERTY GET classTabGroup()
		classTabGroup = "TabGroup"
	END PROPERTY
	
	PROPERTY GET classTabGroupInverted()
		classTabGroup = "TabGroupInverted"
	END PROPERTY
	
	PROPERTY GET classTabGroupVert()
		classTabGroupVert = "TabGroupVert"
	END PROPERTY
	
	PROPERTY GET classTab()
		classTab = "xxxxxx"
	END PROPERTY
	
	PROPERTY GET classSelected()
		classSelected = "SelectedTab"
	END PROPERTY
	
	PROPERTY GET colorTabSelected()
		colorTabSelected = "#FFFFFF"
	END PROPERTY
	
	PROPERTY GET alignTabHoriz()
		alignTabHoriz = "center"
	END PROPERTY
	
	PROPERTY GET imageTL()
		imageTL = "../images/pie_tl_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageTR()
		imageTR = "../images/pie_tr_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageBL()
		imageBL = "../images/pie_bl_gray.gif"
	END PROPERTY

	PROPERTY GET imageBR()
		imageBR = "../images/pie_br_gray.gif"
	END PROPERTY
	
END CLASS


CLASS CTabDataDummy

	PRIVATE m_aData
	PRIVATE m_i
	PRIVATE m_sURL

	PRIVATE SUB Class_Initialize
		m_sURL = ""
		m_i = 0
		m_aData = ARRAY( _
			"One", _
			"Two", _
			"Three", _
			"Four", _
			"Five", _
			"Six", _
			"Seven", _
			"Eight", _
			"Nine", _
			"Ten" _
			)
	END SUB

	
	PROPERTY GET RecordCount()
		RecordCount = UBOUND(m_aData) + 1
	END PROPERTY
	
	PROPERTY GET EOF()
		IF m_i <= UBOUND(m_aData) THEN
			EOF = FALSE
		ELSE
			EOF = TRUE
		END IF
	END PROPERTY
	
	SUB MoveFirst()
		m_i = 0
	END SUB
	
	SUB MoveNext()
		m_i = m_i + 1
	END SUB
	
	FUNCTION IsCurrent( x )
		IF LCASE(x) = LCASE(m_aData(m_i)) THEN
			IsCurrent = TRUE
		ELSE
			IsCurrent = FALSE
		END IF
	END FUNCTION
	
	PROPERTY GET HREF()
		HREF = m_sURL & m_aData(m_i)
	END PROPERTY
	
	PROPERTY GET TabLabel()
		TabLabel = m_aData(m_i)
	END PROPERTY
	
	'----------------
	
	PROPERTY LET URL( s )
		m_sURL = s
	END PROPERTY
	
END CLASS




CLASS CTabGenerate

	PRIVATE m_oData
	PRIVATE m_oFormat
	PRIVATE m_xCurrent
	PRIVATE m_nMaxCols
	
	PRIVATE SUB Class_Initialize
		SET m_oData = Nothing
		SET m_oFormat = Nothing
		m_xCurrent = ""
		m_nMaxCols = 8
	END SUB
	
	PRIVATE SUB Class_Terminate
		SET m_oData = Nothing
		SET m_oFormat = Nothing
	END SUB
	
	PROPERTY SET TabData( o )
		SET m_oData = o
	END PROPERTY
	
	PROPERTY SET TabFormat( o )
		SET m_oFormat = o
	END PROPERTY
	
	PROPERTY LET Current( x )
		m_xCurrent = x
	END PROPERTY
	
	PROPERTY LET MaxCols( n )
		m_nMaxCols = n
	END PROPERTY
	
	PUBLIC SUB makeTabs()
		IF m_nMaxCols < m_oData.RecordCount THEN
			outputTabsVert
		ELSE
			outputTabsHoriz
		END IF
	END SUB
	
	PUBLIC SUB makeTabsInverted()
		IF m_nMaxCols < m_oData.RecordCount THEN
			'outputTabsVert
		ELSE
			outputTabsInvertedHoriz
		END IF
	END SUB
	


	PRIVATE SUB outputTabsGroup( sTabGroup )
	
		DIM	sTBGColor
		DIM	sTabColor
		DIM	sSelectColor
		
		
		DIM	sTabClass
		DIM	sSelectClass
		
		sSelectClass = m_oFormat.classSelected
		
	%>
			<div class="<%=sTabGroup%>">
			<ul>
	<%
	
		IF 0 < m_oData.RecordCount THEN
	
			DIM	sBGColor
			DIM	bThisID
			DIM	sClass
			
			m_oData.MoveFirst
			
			DO UNTIL m_oData.EOF
				IF m_oData.IsCurrent( m_xCurrent ) THEN
					bThisID = TRUE
					sClass = sSelectClass
				ELSE
					bThisID = FALSE
					sClass = ""
				END IF
				Response.Write "<li"
				IF "" <> sClass THEN Response.Write " class=""" & sClass & """"
				Response.Write ">"
				IF "" <> TRIM(m_oData.TabLabel) THEN
				
					Response.Write "<a href=""" & m_oData.HREF & """"
					'IF "" <> sClass THEN Response.Write " class=""" & sClass & """"
					Response.Write ">"
					Response.Write "<span>"
					Response.Write Server.HTMLEncode(m_oData.TabLabel)
					Response.Write "</span>"
					Response.Write "</a>"
				
				ELSE
					Response.Write "&nbsp;"
				END IF
				Response.Write "</li>" & vbCRLF
					
				m_oData.MoveNext
			LOOP
			
		END IF
	%>
		</ul>
		</div>
	<%
	END SUB
	
	
	PRIVATE SUB outputTabsHoriz()
	
		outputTabsGroup m_oFormat.classTabGroup
	
	END SUB
	
	PRIVATE SUB outputTabsInvertedHoriz()
	
		outputTabsGroup m_oFormat.classTabGroupInverted
	
	END SUB
	
	
	SUB outputTabsVert()
	
		outputTabsGroup m_oFormat.classTabGroupVert
	
	END SUB
	

END CLASS


%>