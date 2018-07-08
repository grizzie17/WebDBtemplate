			<table width="175" class="shopmenu" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<th bgcolor="#F3E8CC" class="ShopOnline"><a href="list.asp">Shop Online</a></th>
				</tr>
<%

SUB listOneBrand()
	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"Brands.BrandID, " _
		&		"Brands.Name, " _
		&		"Brands.Logo, " _
		&		"Brands.Disabled " _
		&	"FROM " _
		&		"Brands " _
		&	"WHERE " _
		&		g_sBrand & " = Brands.BrandID " _
		&	"ORDER BY " _
		&		"IF( Brands.Rating IS NULL,99,IF(0=Brands.Rating,99,Brands.Rating)), " _
		&		"Brands.Name " _
		&	";"
		
	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	IF 0 < oRS.RecordCount THEN
	
	
		DIM	oName
		DIM	oID
		DIM	oLogo

		SET oName = oRS("Name")
		SET oID = oRS("BrandID")
		SET oLogo = oRS("Logo")
		
		DIM	sTabName
		sTabName = ""
		
		DIM	n
		n = 0
		
		Response.Write "<tr>"
		Response.Write "<th align=""center"" style=""background-color: #FFFFFF"">"
		
		DO UNTIL oRS.EOF
			n = n + 1
			IF 10 < n THEN EXIT DO
			Response.Write "<a href=""list.asp?brand=" & oID & """>"
			IF "" <> oLogo.Value THEN
				Response.Write "<img border=""0"" src=""picture.asp?table=Brands&name=" & oLogo.Value & """ " _
								& "alt=""" & Server.HTMLEncode( oName.Value ) & """>"
			ELSE
				Response.Write "<i>" & Server.HTMLEncode(oName.Value) & "</i>"
			END IF
			Response.Write "</a>"
			Response.Write "<br>" & vbCRLF
			oRS.MoveNext
		LOOP
		
		Response.Write "</th>"
		Response.Write "</tr>" & vbCRLF

	END IF
	oRS.Close
	SET oRS = Nothing

END SUB

IF "" <> CSTR(g_sBrand) THEN listOneBrand



%>
				<tr>
					<td>
<%



SUB listAllCategories()

	DIM	sBrandID
	DIM	sBrandWhere
	
	sBrandID = Request("brand")
	IF "" <> sBrandID THEN
		sBrandWhere = "AND 0 < (SELECT COUNT(*) FROM Products WHERE " & sBrandID & " = Products.Brand AND Products.Category = Categories.CategoryID) "
	ELSE
		sBrandWhere = ""
	END IF
	
	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"Tabs.TabID AS TabID, " _
		&		"Tabs.Name AS TabName, " _
		&		"Categories.CategoryID AS CatID, " _
		&		"Categories.Name AS CatName, " _
		&		"Categories.Disabled AS CatDisabled " _
		&	"FROM " _
		&		"(TabCategoryMap " _
		&		"INNER JOIN Categories ON (TabCategoryMap.CategoryID = Categories.CategoryID)) " _
		&		"INNER JOIN Tabs ON (TabCategoryMap.TabID = Tabs.TabID) " _
		&	"WHERE " _
		&		"0 = Categories.Disabled " _
		&		"AND 0 = Tabs.Disabled " _
		&		"AND 0 < (SELECT COUNT(*) FROM Products WHERE Products.Category = Categories.CategoryID AND 0 = Products.Disabled) " _
		&		sBrandWhere _
		&	"ORDER BY " _
		&		"IF( Tabs.SortName IS NULL,'',Tabs.SortName)+Tabs.Name, " _
		&		"IF( Categories.SortName IS NULL,'',Categories.SortName)+Categories.Name " _
		&	";"

	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	IF 0 < oRS.RecordCount THEN
	
		DIM	oTabName
		DIM	oTabID
		DIM	oCatName
		DIM	oCatID
		
		SET oTabID = oRS("TabID")
		SET oTabName = oRS("TabName")
		SET oCatID = oRS("CatID")
		SET oCatName = oRS("CatName")
		
		DIM	sTabName
		sTabName = ""
		
		Response.Write "<ul class=""shopcategory"">" & vbCRLF
		DO UNTIL oRS.EOF
			IF sTabName <> oTabName.Value THEN
				'IF "" <> sTabName THEN Response.Write "</ul>" & vbCRLF
				sTabName = oTabName.Value
				Response.Write "<li class=""shoptab"""
				IF "" = g_sQualCategory THEN
					IF "" <> g_sQualTab THEN
						IF g_sQualTab = CSTR(oTabID) THEN
							Response.Write " id=""shopcurrent"""
						END IF
					END IF
				END IF
				Response.Write ">"
				Response.Write "<a href=""list.asp?brand=" & sBrandID & "&tab=" & oTabID.Value & """>"
				Response.Write Server.HTMLEncode( oTabName.Value )
				Response.Write "</a>"
				Response.Write "</li>" & vbCRLF
				'Response.Write "<ul class=""shopcategory"">" & vbCRLF
			END IF
			Response.Write "<li"
			IF "" <> g_sQualCategory THEN
				IF g_sQualCategory = CSTR(oCatID) THEN
					Response.Write " id=""shopcurrent"""
				END IF
			END IF
			Response.Write ">"
			Response.Write "<a href=""list.asp?brand=" & sBrandID & "&tab=" & oTabID.Value & "&category=" & oCatID.Value & """>"
			Response.Write Server.HTMLEncode( oCatName.Value )
			Response.Write "</a>"
			Response.Write "</li>" & vbCRLF
			oRS.MoveNext
		LOOP
		'Response.Write "</ul>" & vbCRLF
		Response.Write "</ul>" & vbCRLF
	
	ELSE
	
		
		Response.Write "<ul class=""shopcategory"">" & vbCRLF
		Response.Write "<li class=""shoptab"">"
		Response.Write "No matching categories"
		Response.Write "</li>" & vbCRLF
		Response.Write "</ul>" & vbCRLF

	END IF
	oRS.Close
	SET oRS = Nothing

END SUB

listAllCategories

IF false THEN
%>
					<font color="#EBCE97">Clubs</font><p>Vintage Collection
						</p>
						<p>Complete Sets</p>
						<p>Drivers</p>
						<p>Fairway Woods</p>
						<p>Hybrids</p>
						<p>Iron Sets</p>
						<p>Individual Irons</p>
						<p>Wedges</p>
						<p>Putters</p>
						<p>&nbsp;</p>
						<p><font color="#EBCE97">Accessories</font></p>
						<p>Apparel &amp; Shoes</p>
						<p>Bags &amp; Accessories</p>
						<p>Push Carts</p>
						<p>Golf Gloves</p>
						<p>Golf Balls</p>
						<p>Training Aids</p>
						<p>Miscellaneous</p>
<%
END IF
%>
						<br>&nbsp;
					</td>
				</tr>
<%


SUB listAllBrands()
	DIM	sCatWhere
	sCatWhere = ""
	
	IF "" <> g_sQualTab THEN
		sCatWhere = sCatWhere & "AND " & g_sQualTab & " = Tabs.TabID "
	END IF
	IF "" <> g_sQualCategory THEN
		sCatWhere = sCatWhere & "AND " & g_sQualCategory & " = Categories.CategoryID "
	END IF
	
	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"Brands.BrandID, " _
		&		"Brands.Name, " _
		&		"Brands.Logo, " _
		&		"Brands.Disabled " _
		&	"FROM " _
		&		"Brands " _
		&	"WHERE " _
		&		"0 = Brands.Disabled " _
		&		"AND 0 < (" _
		&			"SELECT COUNT(*) " _
		&			"FROM " _
		&				"((Products " _
		&				"INNER JOIN Categories ON Categories.CategoryID = Products.Category) " _
		&				"INNER JOIN TabCategoryMap ON Categories.CategoryID = TabCategoryMap.CategoryID) " _
		&				"INNER JOIN Tabs ON (TabCategoryMap.TabID = Tabs.TabID) " _
		&			"WHERE " _
		&				"Products.Brand = Brands.BrandID " _
		&				"AND 0 = Products.Disabled " _
		&				"AND 0 = Tabs.Disabled " _
		&				sCatWhere _
		&			") " _
		&	"ORDER BY " _
		&		"IF( Brands.Rating IS NULL,99,IF(0=Brands.Rating,99,Brands.Rating)), " _
		&		"Brands.Name " _
		&	";"

	DIM oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	IF 0 < oRS.RecordCount THEN

%>
				<tr>
					<th bgcolor="#F6EEDD" class="ShopByBrand">Shop By Brand</th>
				</tr>
				<tr>
					<td class="byBrand">
<%
	
		DIM	oName
		DIM	oID
		DIM	oLogo

		SET oName = oRS("Name")
		SET oID = oRS("BrandID")
		SET oLogo = oRS("Logo")
		
		DIM	sTabName
		sTabName = ""
		
		DIM	n
		n = 0
		
		Response.Write "<ul class=""shopbrands"">" & vbCRLF
		IF "" <> g_sBrand THEN
			Response.Write "<li class=""ShowAllBrands""><a href=""list.asp?brand=&tab=" & g_sQualTab & "&category=" & g_sQualCategory & """>"
			Response.Write "Show All Brands"
			Response.Write "</a>"
			Response.Write "</li>" & vbCRLF
		END IF
		DO UNTIL oRS.EOF
			n = n + 1
			'IF 10 < n THEN EXIT DO
			Response.Write "<li"
			IF "" <> g_sBrand THEN
				IF g_sBrand = CSTR(oID) THEN
					Response.Write " id=""shopcurrent"""
				END IF
			END IF
			Response.Write ">"
			Response.Write "<a href=""list.asp?brand=" & oID.Value & "&tab=" & g_sQualTab & "&category=" & g_sQualCategory & """>"
			IF "" <> oLogo.Value THEN
				Response.Write "<img border=""0"" src=""picture.asp?table=Brands&name=" & oLogo.Value & """ " _
								& "alt=""" & Server.HTMLEncode( oName.Value ) & """>"
			ELSE
				Response.Write "<i>" & Server.HTMLEncode(oName.Value) & "</i>"
			END IF
			'Response.Write Server.HTMLEncode( oName.Value )
			Response.Write "</a>"
			Response.Write "</li>" & vbCRLF
			'Response.Write "<a href=""list.asp?brand=" & oID & """>"
			'Response.Write "</a>"
			'Response.Write "<br>" & vbCRLF
			oRS.MoveNext
		LOOP
		Response.Write "</ul>" & vbCRLF

%>
					</td>
				</tr>
<%
	END IF
	oRS.Close
	SET oRS = Nothing

END SUB

listAllBrands


%>
			</table>
