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
<%


picture_createObject




DIM	sCategoryID
sCategoryID = ""

DIM	sQualCategory
DIM	sQualBrand
sQualCategory = Requestor("QualCategory")
sQualBrand = ""



DIM RID
DIM	sID
DIM	sName
DIM	sSortName
DIM	sShortDesc
DIM	sFormat
DIM	sTitle
DIM	sDesc
DIM	sBody
DIM	sCategory
DIM	bDisabled

RID = Requestor("RID")




SUB submitPicture( sFieldName, nMaxWidth, nMaxHeight )

	DIM	sPictureName
	sPictureName = ""
	
	DIM	sPictureFile
	DIM	sPictureDelete
	
	sPictureDelete = Requestor(sFieldName & "Delete")
	sPictureFile = RequestorFile(sFieldName & "File")
	IF "" <> sPictureFile  OR  "ON" = sPictureDelete THEN
		
		'picture_debug g_sPicturePath & "<br>"
		picture_delete recString( g_RS.Fields(sFieldName) )
		
		IF "" <> sPictureFile THEN
	
			sPictureName = RID & "-" & sFieldName
			sPictureName = picture_saveFile( sFieldName & "File", sPictureFile, sPictureName, nMaxWidth, nMaxHeight )
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
	DIM	sPictureMaxW
	DIM	sPictureMaxH
	DIM	nMaxW
	DIM	nMaxH
	DIM	sLabel
	DIM	nImageID
	
	DIM	n
	FOR n = 1 TO 5
		sLabel = Requestor("Label" & n)
		sPictureFile = RequestorFile("LabelPictureFile" & n)
		IF "" <> sPictureFile  AND  "" <> sLabel THEN
		
			sPictureMaxW = Requestor("LabelPictureMaxW" & n)
			IF ISNUMERIC(sPictureMaxW) THEN
				nMaxW = CLNG(sPictureMaxW)
				IF nMaxW < 1  OR  640 < nMaxW THEN nMaxW = 640
			ELSE
				nMaxW = 640
			END IF
			sPictureMaxH = Requestor("LabelPictureMaxH" & n)
			IF ISNUMERIC(sPictureMaxH) THEN
				nMaxH = CLNG(sPictureMaxH)
				IF nMaxH < 1  OR  640 < nMaxH THEN nMaxH = 640
			ELSE
				nMaxH = 640
			END IF

			sSelect = "" _
				&	"SELECT " _
				&		"* " _
				&	"FROM " _
				&		"pictures " _
				&	"WHERE " _
				&		"PageID = " & RID & " " _
				&		"AND Label = '" & sLabel & "' " _
				&	";"

			SET oRS = dbQueryUpdate( g_DC, sSelect )
			IF 0 < oRS.RecordCount THEN
			'	picture_delete recString( oRS.Fields("File") )
				oRS.Delete 1
			END IF
		
			IF "" <> sPictureFile THEN
	
				sPictureName = RID & "-Library-" & sLabel
				sPictureName = picture_saveFile( "LabelPictureFile" & n, sPictureFile, sPictureName, nMaxW, nMaxH )
				'picture_debug sPictureName & "<br>"
				sPictureFile = picture_buildPath( "pages", sPictureName )
				nImageID = storeImage( g_DC, sPictureFile )

				oRS.AddNew
				
				oRS.Fields("PageID") = RID
				oRS.Fields("Label") = sLabel
				oRS.Fields("ImageID") = nImageID
				oRS.Update

			END IF
		END IF
	NEXT 'n

END SUB

	
		


picture_findPath "Pages"




DIM	sLabelPictureDeleteCount
sLabelPictureDeleteCount = Requestor("LabelPictureDeleteCount")
DIM n
IF ISNUMERIC(sLabelPictureDeleteCount) THEN
	n = CLNG(sLabelPictureDeleteCount)
ELSE
	n = 0
END IF
DIM	sLabelPictureDelete
sLabelPictureDelete = ""
DIM	i
DIM	sTemp
FOR i = 1 TO n
	sTemp = Requestor("LabelPictureDelete" & i)
	IF ISNUMERIC(sTemp) THEN
		sLabelPictureDelete = sLabelPictureDelete & "," & sTemp
	END IF
NEXT
sLabelPictureDelete = MID(sLabelPictureDelete,2)
IF "" <> sLabelPictureDelete THEN

	DIM	sWhere
	IF 0 < INSTR(sLabelPictureDelete, ",") THEN
		sWhere = "PictureID IN (" & sLabelPictureDelete & ") "
	ELSE
		sWhere = "PictureID = " & sLabelPictureDelete & " "
	END IF
	DIM sSelect
	sSelect = "SELECT " _
			&		"PictureID, " _
			&		"PageID, " _
			&		"Label, " _
			&		"ImageID " _
			&	"FROM " _
			&		"pictures " _
			&	"WHERE " _
			&		sWhere  _
			&		"AND pictures.PageID = " & RID _
			&	";"
	
	DIM g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )
	IF 0 < g_RS.RecordCount THEN
		'DIM	sFile
		'DIM	oFile
		'SET oFile = g_RS.Fields("File")
		g_RS.MoveFirst
		DO UNTIL g_RS.EOF
			'sFile = recString( oFile )
			'IF "" <> sFile THEN picture_delete sFile
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

Response.Redirect "images.asp?category=" & sQualCategory

%>