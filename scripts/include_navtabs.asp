<%


DIM g_sNavPageTitle
g_sNavPageTitle = ""


DIM	g_sTopNavTabs
g_sTopNavTabs = ""

DIM	g_nTabSortID
g_nTabSortID = 0

DIM	g_nTabID
g_nTabID = 0

DIM	g_sTabName
g_sTabName = ""

DIM	g_nPageID


DIM	g_sTabSortDetails
g_sTabSortDetails = "ISNULL(categories.SortName,'')+categories.Name, ISNULL(pages.SortName,'')+pages.PageName, pages.Title"
g_sTabSortDetails = "ISNULL(pages.Sortname,'~'), CASE WHEN (schedules.DateEvent IS NULL) THEN DATEDIFF(day,GETDATE(),pages.DateModified) ELSE dbo.LEAST(DATEDIFF(day,schedules.DateEvent,GETDATE()),DATEDIFF(day,GETDATE(),pages.DateModified)) END, pages.PageName "
	'IIF(ISNULL(Pages.Sortname),FORMAT(IIF(ISNULL(Schedules.DateEvent),IIF(ISNULL(Schedules.DateBegin),IIIF(DATEDIFF('d',DATE(),Pages.DateCreated) < 7,DATEADD('d', -50*(7-DATEDIFF('d',DATE(),Pages.DateCreated)),Pages.DateCreated),DATEADD('d',3*DATEDIFF('d',Pages.DateCreated,DATE()),Pages.DateCreated)),IIIF(DATEDIFF('d',DATE(),Schedules.DateBegin) < 7,DATEADD('d', -50*(7-DATEDIFF('d',DATE(),Schedules.DateBegin)),Schedules.DateBegin),DATEADD('d',3*DATEDIFF('d',Schedules.DateBegin,DATE()),Schedules.DateBegin))),IIIF(DATE() <= Schedules.DateEvent,IIIF(DATEDIFF('d',DATE(),Schedules.DateEvent) < 7,DATEADD('d', -100*(7-DATEDIFF('d',DATE(),Schedules.DateEvent)),Schedules.DateEvent),IIIF(ISNULL(Schedules.DateBegin),Schedules.DateEvent,IIF(DATEDIFF('d',Schedules.DateBegin,DATE()) < 8,DATEADD('d',2*(7-DATEDIFF('d',DATE(),Schedules.DateEvent)),Schedules.DateEvent),Schedules.DateEvent))),Schedules.DateEvent)),'yyyy/mm/dd-hh:nn:ss'),Pages.SortName)&Pages.PageName

' all arguments are output
SUB navtabs_gen( sData, sASP, nTab )

	sData = ""
	sASP = ""
	nTab = 0

	DIM	nReqTab
	nReqTab = Request("tab")
	IF NOT ISNUMERIC(nReqTab) THEN
		nReqTab = 0
	ELSE
		nReqTab = CLNG(nReqTab)
		g_nTabID = nReqTab
	END IF
	
	DIM	i
	DIM	sFile
	sFile = ""
	
	IF 0 = nReqTab THEN
		sFile = Request.ServerVariables("SCRIPT_NAME")
		i = INSTRREV( sFile, "/" )
		IF 0 < i THEN
			sFile = LCASE(MID(sFile,i+1))
		ELSE
			i = INSTRREV( sFile, "\" )
			IF 0 < i THEN
				sFile = LCASE(MID(sFile,i+1))
			END IF
		END IF
	END IF


	DIM	xQ
	xQ = """"
	DIM	xQQ
	xQQ = xQ & xQ
	DIM	xNL
	xNL = " & vbCRLF _" & vbCRLF & " & "
	DIM	sOut
	sOut = ""

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"tabs.TabID AS TabID, " _
		&		"tabs.Name AS TabName, " _
		&		"tabs.Description AS TabDesc, " _
		&		"categories.CategoryID AS CatID, " _
		&		"categories.Name AS CatName, " _
		&		"categories.Disabled AS CatDisabled, " _
		&		"pagelistmap.ListID AS TabSortID, " _
		&		"pagelistoptions.Details AS TabSortDetails " _
		&	"FROM " _
		&		"(((tabcategorymap " _
		&		"INNER JOIN categories ON tabcategorymap.CategoryID = categories.CategoryID) " _
		&		"INNER JOIN tabs ON tabcategorymap.TabID = tabs.TabID) " _
		&		"LEFT JOIN pagelistmap ON tabcategorymap.TabID = pagelistmap.TabID) " _
		&		"LEFT JOIN pagelistoptions ON pagelistmap.ListID = pagelistoptions.RID " _
		&	"WHERE " _
		&		"0 = categories.Disabled " _
		&		"AND 0 = tabs.Disabled " _
		&		"AND 0 < (SELECT COUNT(*) FROM pages WHERE pages.Category = categories.CategoryID AND 0 = pages.Disabled) " _
		&	"ORDER BY " _
		&		"ISNULL(tabs.SortName,'')+tabs.Name " _
		&	";"

	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	IF 0 < oRS.RecordCount THEN
	
		DIM	oTabName
		DIM	oTabID
		DIM	oTabDesc
		DIM	oCatName
		DIM	oCatID
		DIM	oTabSortID
		DIM	oTabSortDet
		
		SET oTabID = oRS("TabID")
		SET oTabName = oRS("TabName")
		SET oTabDesc = oRS("TabDesc")
		SET oCatID = oRS("CatID")
		SET oCatName = oRS("CatName")
		SET oTabSortID = oRS("TabSortID")
		SET oTabSortDet = oRS("TabSortDetails")
		
		DIM	nTabID
		DIM	sTabName
		DIM	sTabDesc
		DIM	nTabSortID
		DIM	sTabSortDetails
		nTabID = 0
		sTabName = ""
		sTabDesc = ""
		
		sOut = sOut & xQ & "<div class=" & xQQ & "TopNavigation" & xQQ & ">" & xQ & xNL
		sOut = sOut & xQ & "<ul>" & xQ & xNL
		DO UNTIL oRS.EOF
			IF nTabID <> oTabID.Value THEN
				nTabID = recNumber(oTabID)
				sTabName = recString(oTabName)
				sTabDesc = recString(oTabDesc)
				nTabSortID = recNumber(oTabSortID)
				sTabSortDetails = recString(oTabSortDet)
				sOut = sOut & xQ & "<li "
				IF nTabID = nReqTab THEN
					nTab = nTabID
					sASP = ""
					sData = sData & "g_nTabID = " & nTabID & vbCRLF
					sData = sData & "g_sNavPageTitle = " & xQ & Server.HTMLEncode(sTabName) & xQ & vbCRLF
					sData = sData & "g_sTabName = " & xQ & Server.HTMLEncode(sTabName) & xQ & vbCRLF
					sOut = sOut & "class=" & xQQ & "SelectedTab" & xQQ
					IF 0 < nTabSortID THEN
						sData = sData & "g_nTabSortID = " & nTabSortID & vbCRLF
						sData = sData & "g_sTabSortDetails = " & xQ & sTabSortDetails & xQ & vbCRLF
					END IF
				ELSEIF "@" = LEFT(sTabDesc,1) AND LCASE(MID(sTabDesc,2)) = sFile THEN
					sASP = sFile
					nTab = 0
					sData = sData & "g_nTabID = " & nTabID & vbCRLF
					sData = sData & "g_sNavPageTitle = " & xQ & Server.HTMLEncode(sTabName) & xQ & vbCRLF
					sData = sData & "g_sTabName = " & xQ & Server.HTMLEncode(sTabName) & xQ & vbCRLF
					sOut = sOut &  "class=" & xQQ & "SelectedTab" & xQQ
					IF 0 < nTabSortID THEN
						sData = sData & "g_nTabSortID = " & nTabSortID & vbCRLF
						sData = sData & "g_sTabSortDetails = " & xQ & sTabSortDetails & xQ & vbCRLF
					END IF
				END IF
				sOut = sOut & ">"
				IF "@" = LEFT(sTabDesc,1) THEN
					sOut = sOut & "<a href=" & xQQ & MID(sTabDesc,2) & xQQ & ">"
				ELSE
					sOut = sOut & "<a href=" & xQQ & "page.asp?tab=" & oTabID.Value & xQQ & ">"
				END IF
				sOut = sOut & "<span>" & Server.HTMLEncode( sTabName )
				sOut = sOut & "</span></a>"
				sOut = sOut & "</li>" & xQ & xNL
			END IF
			oRS.MoveNext
		LOOP
		sOut = sOut & xQ & "</ul>" & xQ & xNL
		sOut = sOut & xQ & "</div>" & xQ & vbCRLF
	
	ELSE
	
		
		sOut = sOut & xQ & "<ul class=" & xQQ & "TopNavigation" & xQQ & ">" & xQ & xNL
		sOut = sOut & xQ & "<li class=" & xQQ & "shoptab" & xQQ & ">"
		sOut = sOut &  "No matching categories"
		sOut = sOut & "</li>" & xQ & xNL
		sOut = sOut & xQ & "</ul>" & xQ & vbCRLF

	END IF
	oRS.Close
	SET oRS = Nothing
	
	sData = sData & "g_sTopNavTabs = " & sOut

END SUB



SUB listTabs()


	DIM	sData
	DIM	nTab
	DIM	sASP



	DIM	nReqTab
	nReqTab = Request("tab")
	IF NOT ISNUMERIC(nReqTab) THEN
		nReqTab = 0
	ELSE
		nReqTab = CLNG(nReqTab)
	END IF
	
	DIM	i
	DIM	sFile
	sFile = ""
	
	IF 0 = nReqTab THEN
		sFile = Request.ServerVariables("SCRIPT_NAME")
		i = INSTRREV( sFile, "/" )
		IF 0 < i THEN
			sFile = LCASE(MID(sFile,i+1))
		ELSE
			i = INSTRREV( sFile, "\" )
			IF 0 < i THEN
				sFile = LCASE(MID(sFile,i+1))
			END IF
		END IF
	END IF
	
	DIM	sCacheFile
	IF 0 < nReqTab THEN
		sCacheFile = "tabs-id-" & nReqTab & ".asp"
	ELSE
		sCacheFile = "tabs-file-" & sFile & ".asp"
	END IF
	
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, "d", 28, "m" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE

		navtabs_gen sData, sASP, nTab
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", sCacheFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF

	END IF
	
	EXECUTEGLOBAL sData




END SUB

listTabs


%>