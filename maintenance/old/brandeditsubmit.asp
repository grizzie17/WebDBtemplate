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
	IF "ON" = o THEN
		isTrue = -1
	END IF
END FUNCTION


picture_createObject



DIM	sID
DIM	sName
DIM	sDesc
DIM	sWebsite
DIM	sRating
DIM	bDisabled

sID = Requestor("BrandID")
sName = Requestor("Name")
sDesc = Requestor("Desc")
sWebsite = Requestor("Website")
sRating = Requestor("Rating")
bDisabled = Requestor("Disabled")




		DIM	sSelect
		sSelect = "SELECT " _
			&	" * " _
			&	"FROM Brands " _
			&	"WHERE " _
			&	" BrandID=" & sID & " " _
			&	";"

DIM	g_RS
SET g_RS = dbQueryUpdate( g_DC, sSelect )

	
		
		IF NOT g_RS.EOF THEN
			
			g_RS.Fields("Name").Value = sName
			g_RS.Fields("Description").Value = sDesc
			g_RS.Fields("Website").Value = sWebsite
			g_RS.Fields("Rating").Value = fieldNumber(sRating)
			g_RS.Fields("Disabled").Value = isTrue(bDisabled)


			DIM	sPicturePath
			DIM	sPictureName
			sPictureName = ""
	
			DIM	sPictureFile
			DIM	sPictureDelete
	
			sPictureDelete = Requestor("PictureDelete")
			sPictureFile = Requestor("PictureFile")
			IF "" <> sPictureFile  OR  "ON" = sPictureDelete THEN
		
				picture_findPath "Brands"
				picture_debug g_sPicturePath & "<br>"
				picture_delete recString( g_RS.Fields("Logo") )
		
				IF "" <> sPictureFile THEN
	
					sPictureName = sID & "-Logo"
					sPictureName = picture_saveFile( "PictureFile", sPictureFile, sPictureName )
					picture_debug sPictureName & "<br>"

				END IF
				IF "" = sPictureName THEN sPictureName = NULL
				g_RS.Fields("Logo").Value = sPictureName
			END IF


			g_RS.Update
		
		END IF

		SET g_RS.ActiveConnection = Nothing
		g_RS.Close


SET g_RS = Nothing

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "brands.asp"

%>