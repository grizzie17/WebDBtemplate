<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Server.ScriptTimeout = 60*30

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
DIM	sTitle
DIM	sDesc
DIM	sBody
DIM	sCategory
DIM	sPictureSuffixes
DIM	sDisabled

DIM	sDateBegin
DIM	sDateEnd
DIM	sDateEvent

DIM	bUpdate
	bUpdate = false

RID = Requestor("RID")
sName = Requestor("PageName")
sSortName = Requestor("SortName")
sFormat = Requestor("Format")
sTitle = Requestor("Title")
sDesc = Requestor("Description")
sBody = Requestor("Body")
sCategory = Requestor("Category")
sPictureSuffixes = Requestor("PictureSuffixes")
sDisabled = Requestor("Disabled")

sDateBegin = Requestor("DateBegin")
sDateEnd = Requestor("DateEnd")
sDateEvent = Requestor("DateEvent")

IF "" <> TRIM(sPictureSuffixes) THEN
	g_sPictureSuffixes = TRIM(sPictureSuffixes)
END IF

IF "HTML" = sFormat  AND  0 < LEN(sBody) THEN
	DIM	j
	j = INSTR(sBody, "<body")
	IF j < 1 THEN
		sBody = "<html><body>" & sBody & "</body></html>"
	END IF
END IF


DIM	sSelect
sSelect = "SELECT " _
	&	" * " _
	&	"FROM pages " _
	&	"WHERE " _
	&	" RID=" & RID & " " _
	&	";"

DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )




SUB submitPicture( sFieldName )

	DIM	sPictureName
	sPictureName = ""
	DIM	nImageID
	nImageID = 0
	
	DIM	sPictureFile
	DIM	sPictureMaxW
	DIM	sPictureMAxH
	DIM	nMaxW
	DIM	nMaxH
	DIM	sPictureDelete
	
	sPictureDelete = Requestor(sFieldName & "Delete")
	sPictureFile = RequestorFile(sFieldName & "File")
	IF "" <> sPictureFile  OR  "ON" = sPictureDelete THEN

		sPictureMaxW = Requestor("PictureMaxW")
		IF ISNUMERIC(sPictureMaxW) THEN
			nMaxW = CLNG(sPictureMaxW)
			IF nMaxW < 1  OR  640 < nMaxW THEN nMaxW = 640
		ELSE
			nMaxW = 640
		END IF
		sPictureMaxH = Requestor("PictureMaxH")
		IF ISNUMERIC(sPictureMaxH) THEN
			nMaxH = CLNG(sPictureMaxH)
			IF nMaxH < 1  OR  640 < nMaxH THEN nMaxH = 640
		ELSE
			nMaxH = 640
		END IF

		'picture_debug g_sPicturePath & "<br>"
		picture_delete recString( g_RS.Fields(sFieldName) )
		
		IF "" <> sPictureFile THEN
	
			sPictureName = RID & "-" & sFieldName
			sPictureName = picture_saveFile( sFieldName & "File", sPictureFile, sPictureName, nMaxW, nMaxH )

			sPictureFile = picture_buildPath( "pages", sPictureName )
			nImageID = storeImage( g_DC, sPictureFile )

			'picture_debug sPictureName & "<br>"

		END IF
		g_RS.Fields(sFieldName).Value = nImageID
		bUpdate = true
	END IF

END SUB

	
SUB submitLabelPictures()

	DIM	sPictureName
	sPictureName = ""
	
	DIM	sSelect
	DIM	oRS
	
	DIM	nImageID
	DIM	sPictureFile
	DIM	sPictureMaxW
	DIM	sPictureMaxH
	DIM	nMaxW
	DIM	nMaxH
	DIM	sLabel
	
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
				picture_delete recString( oRS.Fields("File") )
				oRS.Delete 1
			END IF
		
			IF "" <> sPictureFile THEN
	
				sPictureName = RID & "-Label-" & REPLACE(sLabel,".","_")
				sPictureName = picture_saveFile( "LabelPictureFile" & n, sPictureFile, sPictureName, nMaxW, nMaxH )
				'picture_debug sPictureName & "<br>"
				sPictureFile = picture_buildPath( "pages", sPictureName )
				nImageID = storeImage( g_DC, sPictureFile )
				IF 0 <> nImageID THEN
				
					oRS.AddNew
				
					oRS.Fields("PageID") = RID
					oRS.Fields("Label") = sLabel
					oRS.Fields("ImageID") = nImageID
					oRS.Update
				END IF

			END IF
		END IF
	NEXT 'n

END SUB

SUB updateStringField( sField, sValue )
	IF sValue <> recString(g_RS.Fields(sField)) THEN
		IF g_RS.Fields(sField).DefinedSize < LEN(sValue) THEN
			sValue = LEFT(sValue, g_RS.Fields(sField).DefinedSize - 1)
			Response.Write "<p>update Data overflow on field: " & sField & "</p>" & vbCRLF
			Response.Flush
		END IF
		g_RS.Fields(sField).Value = fieldString(sValue)
		bUpdate = true
	END IF
END SUB

SUB updateDateField( sField, dValue )
	IF dValue <> recDate(g_RS.Fields(sField)) THEN
		g_RS.Fields(sField).Value = fieldDate(dValue)
		bUpdate = true
	END IF
END SUB

SUB updateNumberField( sField, nValue )
	IF nValue <> recNumber(g_RS.Fields(sField)) THEN
		g_RS.Fields(sField).Value = fieldNumber(nValue)
		bUpdate = true
	END IF
END SUB

SUB updateBoolField( sField, bValue )
	IF bValue <> recBool(g_RS.Fields(sField)) THEN
		g_RS.Fields(sField).Value = fieldBool(bValue)
		bUpdate = true
	END IF
END SUB

	
		
IF NOT g_RS.EOF THEN

	picture_findPath "pages"
	
	submitPicture "Picture"

	updateStringField "PageName", sName
	updateStringField "SortName", sSortName
	updateStringField "Format", sFormat
	updateStringField "Title", sTitle
	updateStringField "Description", sDesc
	updateStringField "Body", sBody
	updateDateField "DateModified", Now
	updateNumberField "Category", CLNG(sCategory)
	updateBoolField "Disabled", fieldBool(sDisabled)
	IF bUpdate THEN
		g_RS.Update
	END IF
		
END IF

SET g_RS.ActiveConnection = Nothing
g_RS.Close


SET g_RS = Nothing


IF "" = sDateBegin  AND  "" = sDateEnd  AND  "" = sDateEvent  THEN
	' we need to get rid of the Schedules record if it exists


	sSelect = "" _
		&	"SELECT " _
		&	" * " _
		&	"FROM schedules " _
		&	"WHERE " _
		&	" PageID=" & RID & " " _
		&	";"

	SET g_RS = dbQueryUpdate( g_DC, sSelect )
	IF NOT g_RS IS Nothing THEN
		IF NOT g_RS.EOF THEN
			g_RS.MoveFirst
			DO UNTIL g_RS.EOF
				g_RS.Delete 1
				g_RS.MoveNext
			LOOP
		END IF
		g_RS.Close
		SET g_RS = Nothing
	END IF



ELSE


	sSelect = "" _
		&	"SELECT " _
		&	" * " _
		&	"FROM schedules " _
		&	"WHERE " _
		&	" PageID=" & RID & " " _
		&	";"

	SET g_RS = dbQueryUpdate( g_DC, sSelect )
	IF NOT g_RS IS Nothing THEN
		IF NOT g_RS.EOF THEN
		
			g_RS.Fields("DateBegin").Value = fieldDate(sDateBegin)
			g_RS.Fields("DateEnd").Value = fieldDate(sDateEnd)
			g_RS.Fields("DateEvent").Value = fieldDate(sDateEvent)
			g_RS.Fields("Disabled").Value = 0
			g_RS.Update
		
		ELSE
		
			g_RS.AddNew
			g_RS.Fields("PageID").Value = RID
			g_RS.Fields("DateBegin").Value = fieldDate(sDateBegin)
			g_RS.Fields("DateEnd").Value = fieldDate(sDateEnd)
			g_RS.Fields("DateEvent").Value = fieldDate(sDateEvent)
			g_RS.Fields("Disabled").Value = 0
			g_RS.Update

		END IF
		g_RS.Close
		SET g_RS = Nothing
	END IF


END IF





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
			&		sWhere  _
			&		"AND PageID=" & RID & " " _
			&	";"
	
	DIM sImageIDs
	sImageIDs = ""
	SET g_RS = dbQueryUpdate( g_DC, sSelect )
	IF 0 < g_RS.RecordCount THEN
		DIM	sFile
		DIM	nImageID
		DIM	oImageID
		SET oImageID = g_RS.Fields("ImageID")
		DIM	oFile
		SET oFile = g_RS.Fields("File")
		g_RS.MoveFirst
		DO UNTIL g_RS.EOF
			sFile = recString( oFile )
			IF "" <> sFile THEN picture_delete sFile
			'g_RS.Delete 1 'adAffectCurrent
			sImageIDs = sImageIDs & "," & recNumber( oImageID )
			g_RS.MoveNext
		LOOP
	END IF
	g_RS.Close

	sSelect = "DELETE " _
			&	"FROM " _
			&		"pictures " _
			&	"WHERE " _
			&		sWhere  _
			&		"AND PageID=" & RID & " " _
			&	";"
	
	SET g_RS = dbExecute( g_DC, sSelect )
	IF NOTHING IS g_RS THEN
		Response.Write "Problem deleting pictures<br>"
	END IF
	SET g_RS = Nothing

	sImageIDs = MID(sImageIDs,2)
	IF "" <> sImageIDs THEN
		sWhere = "RID IN (" & sImageIDs & ") "
		sSelect = "DELETE " _
				&	"FROM " _
				&		"images " _
				&	"WHERE " _
				&		sWhere  _
				&	";"

		SET g_RS = dbExecute( g_DC, sSelect )
		IF NOTHING IS g_RS THEN
			Response.Write "Problem deleting images<br>"
		END IF
	END IF
	SET g_RS = Nothing


END IF

submitLabelPictures



picture_close

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "pages.asp?category=" & sQualCategory

%>