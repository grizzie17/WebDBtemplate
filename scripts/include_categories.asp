<%

DIM	g_xmlCategories
DIM	g_sTitle

DIM	g_sTab
DIM	g_sTabTitle

DIM	g_sCategory
DIM	g_sCategoryTitle

g_sTitle = ""
g_sTab = ""
g_sTabTitle = ""

g_sCategory = ""
g_sCategoryTitle = ""

g_sTab = Request.QueryString("tab")
g_sCategory = Request.QueryString("category")

'SET g_xmlCategories = Server.CreateObject("msxml2.DOMDocument")


FUNCTION category_getName( oDC, sSelect )
	category_getName = ""
	DIM	oRS
	SET oRS = dbQueryRead( oDC, sSelect )
	IF NOT Nothing IS oRS THEN
		IF 0 < oRS.RecordCount THEN
			category_getName = oRS.Fields("Name").Value
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
END FUNCTION

SUB loadCategoryXML()

	IF "" <> g_sTab  OR  "" <> g_sCategory THEN
		
		DIM	sSelect
		DIM	oRS
			IF "" <> g_sTab THEN
				sSelect = "" _
					&	"SELECT " _
					&		"Name " _
					&	"FROM " _
					&		"tabs " _
					&	"WHERE " _
					&		"TabID = " & g_sTab _
					&	";"
					g_sTabTitle = category_getName( g_DC, sSelect )
					g_sTitle = g_sTabTitle
			END IF
			IF "" <> g_sCategory THEN
				sSelect = "" _
					&	"SELECT " _
					&		"Name " _
					&	"FROM " _
					&		"categories " _
					&	"WHERE " _
					&		"CategoryID = " & g_sCategory _
					&	";"
					g_sCategoryTitle = category_getName( g_DC, sSelect )
			END IF
	END IF
	
END SUB
loadCategoryXML


%>