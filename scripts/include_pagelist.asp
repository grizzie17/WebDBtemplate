<%


' This function assumes g_sTabSortDetails is set
FUNCTION getPageListRS( nTabID )

	SET getPageListRS = Nothing

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"pages.RID, " _
		&		"pages.PageName, " _
		&		"pages.Format, " _
		&		"pages.Title, " _
		&		"pages.Description, " _
		&		"pages.Category, " _
		&		"pages.Picture, " _
		&		"tabcategorymap.TabID, " _
		&		"categories.Name AS CatName, " _
		&		"categories.ShortName AS CatShortName, " _
		&		"schedules.DateBegin, " _
		&		"schedules.DateEnd, " _
		&		"schedules.DateEvent, " _
		&		"pages.DateModified, " _
		&		"ISNULL(schedules.DateBegin,pages.DateCreated) AS dtBegin " _
		&	"FROM " _
		&		"((pages " _
		&		"INNER JOIN categories ON categories.CategoryID = pages.Category) " _
		&		"INNER JOIN tabcategorymap ON categories.CategoryID = tabcategorymap.CategoryID) " _
		&		"LEFT JOIN schedules ON schedules.PageID = pages.RID " _
		&	"WHERE " _
		&		"tabcategorymap.TabID = " & nTabID & " " _
		&		"AND 0 = pages.Disabled " _
		&		"AND (" _
		&		"ISNULL(schedules.DateBegin,GETDATE()) <= GETDATE() " _
		&		" AND " _
		&		"GETDATE() <= ISNULL(schedules.DateEnd,GETDATE()) " _
		&		") " _
		&	"ORDER BY " _
		&		g_sTabSortDetails & " " _
		&	";"
		
	'Response.Write Server.HTMLEncode(sSelect) & "<br>"

	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )

	IF NOT oRS IS Nothing THEN
		IF oRS.EOF THEN
			SET getPageListRS = Nothing
			oRS.Close
		ELSE
			SET getPageListRS = oRS
		END iF
		SET oRS = Nothing
	END IF


END FUNCTION


' This function assumes g_sTabSortDetails is set
FUNCTION getPagesRS( nTabID )

	SET getPagesRS = Nothing

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"pages.RID, " _
		&		"pages.PageName, " _
		&		"pages.Format, " _
		&		"pages.Title, " _
		&		"pages.Description, " _
		&		"pages.Body, " _
		&		"pages.Category, " _
		&		"pages.Picture, " _
		&		"tabcategorymap.TabID, " _
		&		"categories.Name AS CatName, " _
		&		"categories.ShortName AS CatShortName, " _
		&		"schedules.DateBegin, " _
		&		"schedules.DateEnd, " _
		&		"schedules.DateEvent, " _
		&		"pages.DateModified, " _
		&		"ISNULL(schedules.DateBegin,pages.DateCreated) AS dtBegin " _
		&	"FROM " _
		&		"((pages " _
		&		"INNER JOIN categories ON categories.CategoryID = pages.Category) " _
		&		"INNER JOIN tabcategorymap ON categories.CategoryID = tabcategorymap.CategoryID) " _
		&		"LEFT JOIN schedules ON schedules.PageID = pages.RID " _
		&	"WHERE " _
		&		"tabcategorymap.TabID = " & nTabID & " " _
		&		"AND SUBSTRING( categories.Name, 1, 1 ) != '@' " _
		&		"AND 0 = pages.Disabled " _
		&		"AND (" _
		&		"ISNULL(schedules.DateBegin,GETDATE()) <= GETDATE() " _
		&		" AND " _
		&		"GETDATE() <= ISNULL(schedules.DateEnd,GETDATE()) " _
		&		") " _
		&	"ORDER BY " _
		&		g_sTabSortDetails & " " _
		&	";"

	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )

	IF NOT oRS IS Nothing THEN
		IF oRS.EOF THEN
			SET getPagesRS = Nothing
			oRS.Close
		ELSE
			SET getPagesRS = oRS
		END iF
		SET oRS = Nothing
	END IF


END FUNCTION


FUNCTION defineTabSortDetails( sTabName )

	defineTabSortDetails = 0

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"tabs.TabID, " _
		&		"pagelistoptions.Details " _
		&	"FROM " _
		&		"(tabs " _
		&		"LEFT JOIN pagelistmap ON tabs.TabID = pagelistmap.TabID) " _
		&		"LEFT JOIN pagelistoptions ON pagelistmap.ListID = pagelistoptions.RID " _
		&	"WHERE " _
		&		"tabs.Name LIKE '" & sTabName & "' " _
		&	";"

	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		IF NOT oRS.EOF THEN
			defineTabSortDetails = recNumber(oRS.Fields("TabID"))
			
			DIM	sSort
			sSort = recString(oRS.Fields("Details"))
			IF "" <> sSort THEN
				g_sTabSortDetails = sSort
			END IF
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF

END FUNCTION



SUB outputMultiplePages( bOutputTable )

	DIM	oRS
	SET oRS = getPagesRS( g_nTabID )
	IF NOT oRS IS Nothing THEN
	
		IF 0 < oRS.RecordCount THEN

			DIM	oPageRID
			DIM	oFormat
			DIM	oTitle
			DIM	oDescription
			DIM	oBody
			DIM	oPicture
			DIM	oDateModified
			DIM	oCatName

			SET oPageRID = oRS.Fields("RID")
			SET oFormat = oRS.Fields("Format")
			SET oTitle = oRS.Fields("Title")
			SET oDescription = oRS.Fields("Description")
			SET oBody = oRS.Fields("Body")
			SET oPicture = oRS.Fields("Picture")
			SET oDateModified = oRS.Fields("DateModified")
			SET	oCatName = oRS.Fields("CatName")

			DIM	sFormat
			DIM	sTitle
			DIM	sDescription
			DIM	sBody
			DIM	sPicture
			DIM	dDateModified
			DIM	sCatName

			IF bOutputTable THEN
				Response.Write "<table width=""100%"">" & vbCRLF
			END IF

			oRS.MoveFirst
			DO UNTIL oRS.EOF
				SET g_oPictureDict = Nothing
				sCatName = recString(oCatName)
				IF "@" <> LEFT(sCatName,1) THEN
					sFormat = recString(oFormat)
					sTitle = pagebody_formatString( sFormat, recString(oTitle) )
					dDateModified = recDate(oDateModified)
					IF g_dLocalPageDateModified < dDateModified THEN g_dLocalPageDateModified = dDateModified

					g_nPageID = recNumber(oPageRID)
					g_sXDescription = recString(oDescription)
					g_sXBody = recString(oBody)
					g_sXPicture = recString(oPicture)
		
					'Response.Write "<div style=""background-color:#FFDD99; text-align:center; font-weight:bold; width: 100%"">" & sTitle & "</div>" & vbCRLF
					Response.Write "<tr>"
					Response.Write "<th bgcolor=""#FFDD99"" align=""center"" class=""BlockHead"">" & sTitle & "</th>" & vbCRLF
					Response.Write "</tr>" & vbCRLF
					Response.Write "<tr>"
					Response.Write "<td class=""BlockBody"">"
					'Response.Write "<div>"
					outputFormattedPageInfo sFormat, g_sXDescription, g_sXBody, g_sXPicture
					'Response.Write "</div>"
					Response.Write "<div style=""text-align: left; color: #999999; font-family: sans-serif; font-size: xx-small;"">"
					Response.Write "Updated: " & Server.HTMLEncode(DATEADD("h", g_nServerTimeZoneOffset, dDateModified))
					Response.Write "</div>"
					Response.Write "</td>"
					Response.Write "</tr>" & vbCRLF
				END IF
				oRS.MoveNext
			LOOP

			IF bOutputTable THEN
				Response.Write "</table>" & vbCRLF
			END IF
		END IF
	
		oRS.Close
		SET oRS = Nothing
	END IF

END SUB


%>