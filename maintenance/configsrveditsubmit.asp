<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<%



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
		sPictureFile = RequestorFile("LabelPictureFile" & n)
		IF "" <> sPictureFile  AND  "" <> sLabel THEN
		
			sSelect = "" _
				&	"SELECT " _
				&		"* " _
				&	"FROM " _
				&		"Pictures " _
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
	
				sPictureName = RID & "-Label-" & sLabel
				sPictureName = picture_saveFile( "LabelPictureFile" & n, sPictureFile, sPictureName )
				'picture_debug sPictureName & "<br>"
				
				oRS.AddNew
				
				oRS.Fields("PageID") = RID
				oRS.Fields("Label") = sLabel
				oRS.Fields("File") = sPictureName
				oRS.Update

			END IF
		END IF
	NEXT 'n

END SUB








picture_createObject




DIM	sCategoryID
sCategoryID = Requestor("CategoryID")

DIM	sQualCategory
DIM	sQualBrand
sQualCategory = Requestor("QualCategory")



DIM	sRID
DIM	sDomain
DIM	sSiteName
DIM	sSiteTitle
DIM	sSiteOrgDesignation
DIM	sSiteChapter
DIM	sSiteChapterID
DIM	sSiteMotto
DIM	sSiteLocation
DIM	sSiteZip
DIM	sCopyrightStartYear
DIM	sMailServer
DIM	sMailUser
DIM	sMailPW
DIM	sSiteTabAnnouncements
DIM	sSiteTabNewsletters
DIM	sPicture
DIM	sDisabled



sRID = Requestor("RID")
sDomain = Requestor("Domain")
sMailServer = Requestor("MailServer")
sMailUser = Requestor("MailUser")
sMailPW = Requestor("MailPW")





DIM	sSelect
sSelect = "SELECT " _
	&	" * " _
	&	"FROM Config " _
	&	"WHERE " _
	&	" RID=" & sRID & " " _
	&	";"

DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )



	
		
IF NOT g_RS.EOF THEN

	'picture_findPath "Pages"
	
	'submitPicture "Picture", 200, 0




	g_RS.Fields("Domain").Value = fieldString(sDomain)
	g_RS.Fields("MailServer").Value = fieldString(sMailServer)
	g_RS.Fields("MailUser").Value = fieldString(sMailUser)
	g_RS.Fields("MailPW").Value = fieldString(sMailPW)

	g_RS.Update
		
END IF

SET g_RS.ActiveConnection = Nothing
g_RS.Close


SET g_RS = Nothing




IF FALSE THEN ' disable pictures

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
			&		"PageID, " _
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

END IF ' disable pictures


picture_close

cache_clearFolder "site"

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "./"

%>