<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<!--#include file="scripts\include_relations.asp"-->
<%


SUB deleteSupportingPictures( id )
	DIM	sSelect
	sSelect = "SELECT " _
			&		"PictureID, " _
			&		"PageID, " _
			&		"Label, " _
			&		"ImageID, " _
			&		dbQ("File") & " " _
			&	"FROM " _
			&		"(pictures " _
			&		"LEFT JOIN images ON pictures.ImageID = images.RID) " _
			&	"WHERE " _
			&		"PageID=" & id & " " _
			&	";"
	
	DIM sImageIDs
	sImageIDs = ""
	DIM	oRS
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF 0 < oRS.RecordCount THEN
		DIM	sFile
		DIM	nImageID
		DIM	oImageID
		SET oImageID = oRS.Fields("ImageID")
		DIM	oFile
		SET oFile = oRS.Fields("File")
		oRS.MoveFirst
		DO UNTIL oRS.EOF
			sFile = recString( oFile )
			IF "" <> sFile THEN picture_delete sFile
			'g_RS.Delete 1 'adAffectCurrent
			sImageIDs = sImageIDs & "," & recNumber( oImageID )
			oRS.MoveNext
		LOOP
	END IF
	oRS.Close

	sSelect = "DELETE " _
			&	"FROM " _
			&		"pictures " _
			&	"WHERE " _
			&		"PageID=" & id & " " _
			&	";"
	
	SET oRS = dbExecute( g_DC, sSelect )
	IF NOTHING IS oRS THEN
		Response.Write "Problem deleting pictures<br>"
	END IF
	SET oRS = Nothing

	sImageIDs = MID(sImageIDs,2)
	IF "" <> sImageIDs THEN
		DIM sWhere
		sWhere = "RID IN (" & sImageIDs & ") "
		sSelect = "DELETE " _
				&	"FROM " _
				&		"images " _
				&	"WHERE " _
				&		sWhere  _
				&	";"

		SET oRS = dbExecute( g_DC, sSelect )
		IF NOTHING IS oRS THEN
			Response.Write "Problem deleting images<br>"
		END IF
	END IF
	SET oRS = Nothing

END SUB


SUB deleteSchedules( id )

	DIM sSelect
	sSelect = "" _
		&	"SELECT " _
		&	" * " _
		&	"FROM schedules " _
		&	"WHERE " _
		&	" PageID=" & id & " " _
		&	";"

	DIM oRS
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		IF NOT oRS.EOF THEN
			oRS.MoveFirst
			DO UNTIL oRS.EOF
				oRS.Delete 1
				oRS.MoveNext
			LOOP
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF

END SUB


DIM	sID
sID = TRIM(Request("page"))

DIM	sQualCategory
sQualCategory = Request("QualCategory")


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
		&	"FROM pages " _
		&	"WHERE " _
		&		sWhere _
		&	";"

	DIM	g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
	IF NOT g_RS.EOF THEN
	
		DIM sFile
		
		DIM	oPicture
		SET oPicture = g_RS.Fields("Picture")
		
		
		DIM	oID
		SET oID = g_RS.Fields("RID")

		picture_findPath "Pages"
		
		g_RS.MoveFirst		
		DO UNTIL g_RS.EOF
			sFile = recString(oPicture)
			IF "" <> sFile THEN picture_delete sFile

			deleteSupportingPictures recNumber(oID)
			
			deleteSchedules recNumber(oID)
			
			g_RS.Delete adAffectCurrent
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

Response.Redirect "pages.asp?category=" & sQualCategory

%>