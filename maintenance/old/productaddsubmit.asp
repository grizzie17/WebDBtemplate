<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%

FUNCTION isTrue( o )
	isTrue = 0
	IF "ON" = UCASE(o) THEN
		isTrue = -1
	END IF
END FUNCTION


FUNCTION fieldNumber( s )
	IF "" = CSTR(s) THEN
		fieldNumber = NULL
	ELSEIF ISNUMERIC( s ) THEN
		fieldNumber = s
	ELSE
		fieldNumber = 0
	END IF
END FUNCTION


picture_createObject

DIM	RID
DIM	sCategory
DIM	sProdID
DIM	sName
DIM	sSortName
DIM	sShortDesc
DIM	sFormat
DIM	sDesc
DIM	sBrand
DIM	sUnitPrice
DIM	sListPrice
DIM	sMAPrice
DIM	sCost
DIM	sWeight
DIM	sHideQuantity
DIM	sWebsite
DIM	sDateAdded

DIM	sQualCategory
sQualCategory = Requestor("QualCategory")

DIM	sQualBrand
sQualBrand = Requestor("QualBrand")

sCategory = Requestor("Category")
sProdID = Requestor("ProdID")
sName = Requestor("Name")
sSortName = Requestor("SortName")
sShortDesc = Requestor("ShortDesc")
sFormat = Requestor("Format")
sDesc = Requestor("Desc")
sBrand = Requestor("Brand")
sUnitPrice = Requestor("UnitPrice")
sListPrice = Requestor("ListPrice")
sMAPrice = Requestor("MAPrice")
sCost = Requestor("Cost")
sWeight = Requestor("Weight")
sHideQuantity = Requestor("HideQuantity")
sWebsite = Requestor("Website")
sDateAdded = Requestor("DateAdded")

IF "" = TRIM(sDateAdded) THEN
	sDateAdded = NOW
END IF




		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM Products " _
			&	"WHERE " _
			&	" RID=0" _
			&	";"
	
DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )





		g_RS.AddNew
		
		g_RS.Fields("ProdID").Value = sProdID
		g_RS.Fields("Name").Value = sName
		g_RS.Fields("SortName").Value = sSortName
		g_RS.Fields("ShortDesc").Value = sShortDesc
		g_RS.Fields("Format").Value = sFormat
		g_RS.Fields("Description").Value = sDesc
		g_RS.Fields("DateAdded").Value = sDateAdded
		g_RS.Fields("Category").Value = sCategory
		g_RS.Fields("Brand").Value = sBrand
		g_RS.Fields("UnitPrice").Value = fieldNumber(sUnitPrice)
		g_RS.Fields("ListPrice").Value = fieldNumber(sListPrice)
		g_RS.Fields("MAPrice").Value = fieldNumber(sMAPrice)
		g_RS.Fields("Cost").Value = fieldNumber(sCost)
		g_RS.Fields("Weight").Value = fieldNumber(sWeight)
		g_RS.Fields("HideQuantity").Value = isTrue(sHideQuantity)
		g_RS.Fields("URL").Value = sWebsite
		g_RS.Fields("Disabled").Value = 0
		
		g_RS.Update
		RID = g_RS.Fields("RID").Value

		DIM	sPicturePath
		DIM	sPictureName
		sPictureName = ""
		
		DIM	sPictureFile
		DIM	sLargePictureFile
		DIM	sThumbnailFile	
		sPictureFile = Requestor("PictureFile")
		sLargePictureFile = Requestor("LargePictureFile")
		sThumbnailFile = Requestor("ThumbnailFile")
		IF "" <> sPictureFile  OR  "" <> sLargePictureFile  OR  "" <> sThumbnailFile THEN
			
			picture_findPath "Products"
			
			IF "" <> sPictureFile THEN
				sPictureName = RID & "-Picture"
				sPictureName = picture_saveFile( "PictureFile", sPictureFile, sPictureName )
				g_RS.Fields("Picture").Value = sPictureName
			END IF
			IF "" <> sLargePictureFile THEN
				sPictureName = RID & "-LargePicture"
				sPictureName = picture_saveFile( "LargePictureFile", sLargePictureFile, sPictureName )
				g_RS.Fields("LargePicture").Value = sPictureName
			END IF
			IF "" <> sThumbnailFile THEN
				sPictureName = RID & "-Thumbnail"
				sPictureName = picture_saveFile( "ThumbnailFile", sThumbnailFile, sPictureName )
				g_RS.Fields("Thumbnail").Value = sPictureName
			END IF
			
			g_RS.Update
	
		END IF
		
		SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "products.asp?category=" & sQualCategory & "&brand=" & sQualBrand

%>