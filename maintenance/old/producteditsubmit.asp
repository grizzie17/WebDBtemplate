<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%


picture_createObject




DIM	sCategoryID
sCategoryID = Requestor("CategoryID")

DIM	sQualCategory
DIM	sQualBrand
sQualCategory = Requestor("QualCategory")
sQualBrand = Requestor("QualBrand")



DIM RID
DIM	sID
DIM	sName
DIM	sSortName
DIM	sShortDesc
DIM	sFormat
DIM	sDesc
DIM	sCategory
DIM	sBrand
DIM	sUnitPrice
DIM	sListPrice
DIM	sMAPrice
DIM	sCost
DIM	sWeight
DIM	sHideQuantity
DIM	sOptions
DIM	sStockCount
DIM	sDateAdded
DIM	sWebsite
DIM	bDisabled
DIM	sBody

RID = Requestor("RID")
sID = Requestor("ProdID")
sName = Requestor("Name")
sSortName = Requestor("SortName")
sShortDesc = Requestor("ShortDesc")
sFormat = Requestor("Format")
sDesc = Requestor("Desc")
sCategory = Requestor("Category")
sBrand = Requestor("Brand")
sUnitPrice = Requestor("UnitPrice")
sListPrice = Requestor("ListPrice")
sMAPrice = Requestor("MAPrice")
sCost = Requestor("Cost")
sWeight = Requestor("Weight")
sHideQuantity = Requestor("HideQuantity")
sOptions = Requestor("Options")
sStockCount = Requestor("StockCount")
sDateAdded = Requestor("DateAdded")
sWebsite = Requestor("Website")
bDisabled = Requestor("Disabled")
sBody = TRIM(Requestor("Body"))


DIM	sSelect
sSelect = "SELECT " _
	&	" * " _
	&	"FROM Products " _
	&	"WHERE " _
	&	" RID=" & RID & " " _
	&	";"

DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )


SUB submitPicture( sFieldName, nMaxWidth, nMaxHeight )

	DIM	sPictureName
	sPictureName = ""
	
	DIM	sPictureFile
	DIM	sPictureDelete
	
	sPictureDelete = Requestor(sFieldName & "Delete")
	sPictureFile = Requestor(sFieldName & "File")
	IF "" <> sPictureFile  OR  "ON" = sPictureDelete THEN
		
		'picture_debug g_sPicturePath & "<br>"
		picture_delete recString( g_RS.Fields(sFieldName) )
		
		IF "" <> sPictureFile THEN
	
			sPictureName = RID & "-" & sFieldName
			sPictureName = picture_saveFile( sFieldName & "File", sPictureFile, sPictureName )
			'picture_debug sPictureName & "<br>"

		END IF
		g_RS.Fields(sFieldName).Value = sPictureName
	END IF

END SUB

	
SUB submitLabelPictures()

	DIM	sPictureName
	sPictureName = ""
	
	DIM	sSelect
	DIM	oRS
	
	DIM	sPictureFile
	DIM	sLabel
	
	DIM	n
	FOR n = 1 TO 5
		sLabel = Requestor("Label" & n)
		sPictureFile = Requestor("LabelPictureFile" & n)
		IF "" <> sPictureFile  AND  "" <> sLabel THEN
		
			sSelect = "" _
				&	"SELECT " _
				&		"* " _
				&	"FROM " _
				&		"Pictures " _
				&	"WHERE " _
				&		"ProductID = " & RID & " " _
				&		"AND Label = '" & sLabel & "' " _
				&	";"

			SET oRS = dbQueryUpdate( g_DC, sSelect )
			IF 0 < oRS.RecordCount THEN
				picture_delete recString( oRS.Fields("File") )
				oRS.Delete 1
			END IF
		
			IF "" <> sPictureFile THEN
	
				sPictureName = RID & "-Label-" & sLabel
				sPictureName = picture_saveFile( "LabelPictureFile" & n, sPictureFile, sPictureName )
				'picture_debug sPictureName & "<br>"
				
				oRS.AddNew
				
				oRS.Fields("ProductID") = RID
				oRS.Fields("Label") = sLabel
				oRS.Fields("File") = sPictureName
				oRS.Update

			END IF
		END IF
	NEXT 'n

END SUB

	
		
IF NOT g_RS.EOF THEN

	picture_findPath "Products"
	
	submitPicture "Picture", 200, 0
	submitPicture "LargePicture", 500, 500
	submitPicture "Thumbnail", 50, 50

	g_RS.Fields("ProdID").Value = sID
	g_RS.Fields("Name").Value = sName
	g_RS.Fields("SortName").Value = sSortName
	g_RS.Fields("ShortDesc").Value = sShortDesc
	g_RS.Fields("Format").Value = sFormat
	g_RS.Fields("Description").Value = sDesc
	g_RS.Fields("Category").Value = sCategory
	g_RS.Fields("Brand").Value = sBrand
	g_RS.Fields("UnitPrice").Value = fieldNumber(sUnitPrice)
	g_RS.Fields("ListPrice").Value = fieldNumber(sListPrice)
	g_RS.Fields("MAPrice").Value = fieldNumber(sMAPrice)
	g_RS.Fields("Cost").Value = fieldNumber(sCost)
	g_RS.Fields("Weight").Value = fieldNumber(sWeight)
	g_RS.Fields("HideQuantity").Value = fieldBool(sHideQuantity)
	g_RS.Fields("Options").Value = sOptions
	g_RS.Fields("StockCount").Value = fieldNumber(sStockCount)
	g_RS.Fields("DateAdded").Value = fieldDate(sDateAdded)
	g_RS.Fields("URL").Value = sWebsite
	g_RS.Fields("Disabled").Value = fieldBool(bDisabled)
	g_RS.Update
		
END IF

SET g_RS.ActiveConnection = Nothing
g_RS.Close


SET g_RS = Nothing


sSelect = "SELECT " _
	&	" * " _
	&	"FROM Gallery " _
	&	"WHERE " _
	&	" ProductID=" & RID & " " _
	&	";"

SET g_RS = dbQueryUpdate( g_DC, sSelect )

IF "" = sBody THEN
	' need to delete any existing records

	IF 0 < g_RS.RecordCount THEN
		g_RS.MoveFirst
		DO UNTIL g_RS.EOF
			g_RS.Delete 1 'adAffectCurrent
			g_RS.MoveNext
		LOOP
	END IF

ELSE

	IF 0 = g_RS.RecordCount THEN
		g_RS.AddNew
	END IF
	
	g_RS.Fields("ProductID").Value = RID
	g_RS.Fields("Body").Value = sBody
	g_RS.Update
	
END IF

SET g_RS.ActiveConnection = Nothing
g_RS.Close
SET g_RS = Nothing



DIM	sLabelPictureDelete
sLabelPictureDelete = Requestor("LabelPictureDelete")
IF "" <> sLabelPictureDelete THEN

	DIM	sWhere
	IF 0 < INSTR(sLabelPictureDelete, ",") THEN
		sWhere = "PictureID IN (" & sLabelPictureDelete & ") "
	ELSE
		sWhere = "PictureID = " & sLabelPictureDelete & " "
	END IF
	sSelect = "SELECT " _
			&		"PictureID, " _
			&		"ProductID, " _
			&		"Label, " _
			&		"File " _
			&	"FROM " _
			&		"Pictures " _
			&	"WHERE " _
			&		sWhere  _
			&	";"
	
	
	SET g_RS = dbQueryUpdate( g_DC, sSelect )
	IF 0 < g_RS.RecordCount THEN
		DIM	sFile
		DIM	oFile
		SET oFile = g_RS.Fields("File")
		DO UNTIL g_RS.EOF
			sFile = recString( oFile )
			IF "" <> sFile THEN picture_delete sFile
			g_RS.Delete 1 'adAffectCurrent
			g_RS.MoveNext
		LOOP
	END IF
	g_RS.Close

END IF

submitLabelPictures



picture_close

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "products.asp?category=" & sQualCategory & "&brand=" & sQualBrand

%>