<%@ LANGUAGE="VBSCRIPT" %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<!--#include file="scripts\include_relations.asp"-->
<%


SUB deleteSupportingPictures( id )
	DIM	sSelect
	sSelect = "SELECT " _
		&		"* " _
		&	"FROM Pictures " _
		&	"WHERE " _
		&		"ProductID=" & id & " " _
		&	";"
	
	DIM	oRS
	SET oRS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT oRS.EOF THEN
	
		DIM sFile
		
		DIM	oPicture
		SET oPicture = oRS.Fields("File")
		
		oRS.MoveFirst		
		DO UNTIL oRS.EOF
			sFile = recString(oPicture)
			IF "" <> sFile THEN picture_delete sFile
			
			oRS.Delete 1 'adAffectCurrent
			oRS.MoveNext
		LOOP
	END IF
	
	Set oRS.ActiveConnection = Nothing
	oRS.Close
	SET oRS = Nothing

END SUB


DIM	sID
sID = TRIM(Request("product"))

DIM	sQualCategory
sQualCategory = Request("QualCategory")

DIM	sQualBrand
sQualBrand = Request("QualBrand")

IF "" <> sID THEN

	DIM	sWhere
	IF 0 < INSTR(sID,",") THEN
		sWhere = " RID IN (" & sID & ") "
	ELSE
		sWhere = " RID=" & sID & " "
	END IF


	DIM	sSelect
	sSelect = "SELECT " _
		&		"* " _
		&	"FROM Products " _
		&	"WHERE " _
		&		sWhere _
		&	";"

	DIM	g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT g_RS.EOF THEN
	
		DIM sFile
		
		DIM	oPicture
		SET oPicture = g_RS.Fields("Picture")
		
		DIM	oLargePicture
		SET oLargePicture = g_RS.Fields("LargePicture")
		
		DIM	oThumbnail
		SET oThumbnail = g_RS.Fields("Thumbnail")
		
		DIM	oID
		SET oID = g_RS.Fields("RID")

		picture_findPath "Products"
		
		g_RS.MoveFirst		
		DO UNTIL g_RS.EOF
			sFile = recString(oPicture)
			IF "" <> sFile THEN picture_delete sFile
			sFile = recString(oLargePicture)
			IF "" <> sFile THEN picture_delete sFile
			sFile = recString(oThumbnail)
			IF "" <> sFile THEN picture_delete sFile

			deleteSupportingPictures oID.Value
			
			g_RS.Delete 1 'adAffectCurrent
			g_RS.MoveNext
		LOOP

	END IF
	
	Set g_RS.ActiveConnection = Nothing
	g_RS.Close
	SET g_RS = Nothing

END IF

SET g_RS = Nothing

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "products.asp?category=" & sQualCategory & "&brand=" & sQualBrand

%>