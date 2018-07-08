<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_relations.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%

FUNCTION isTrue( o )
	isTrue = fieldBool( o )
END FUNCTION



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

	


g_bUseRequestObject = TRUE

picture_createObject




DIM	sTabID
DIM	sName
DIM	sSortName
DIM	sDescription
DIM	sPageSort
DIM	bDisabled
DIM	sCategories

sTabID = Requestor("TabID")
sName = Requestor("Name")
sSortName = Requestor("SortName")
sDescription = Requestor("ShortDesc")
sPageSort = Requestor("PageSort")
bDisabled = Requestor("Disabled")
sCategories = Requestor("Categories")




		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM tabs " _
			&	"WHERE " _
			&	" TabID=" & sTabID & " " _
			&	";"

DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
		
		IF NOT g_RS.EOF THEN

			'picture_findPath "Tabs"
	
			'submitPicture "Picture"

			
			g_RS.Fields("Name").Value = sName
			g_RS.Fields("SortName").Value = sSortName
			g_RS.Fields("Description").Value = sDescription
			g_RS.Fields("Disabled").Value = isTrue(bDisabled)
			g_RS.Update
		
		END IF

		'SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

commitCategories CLNG(sTabID), sCategories



		sSelect = "SELECT " _
			&	" * " _
			&	"FROM pagelistmap " _
			&	"WHERE " _
			&	" TabID=" & sTabID & " " _
			&	";"

		SET g_RS = dbQueryUpdate( g_DC, sSelect )
		IF NOT g_RS IS Nothing THEN
			IF NOT g_RS.EOF THEN
				IF "" <> sPageSort THEN
					g_RS.Fields("ListID").Value = fieldNumber( sPageSort )
					g_RS.Update
				ELSE
					g_RS.Delete 1
				END IF
			ELSE
				g_RS.AddNew
				g_RS.Fields("TabID").Value = fieldNumber( sTabID )
				g_RS.Fields("ListID").Value = fieldNumber( sPageSort )
				g_RS.Update
			END IF
			g_RS.Close
			SET g_RS = Nothing
		END IF


picture_close

cache_clearFolder "site"


%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "tabs.asp"

%>