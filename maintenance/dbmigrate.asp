<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin_admin.asp"-->
<!--#include file="scripts\include_dbaccess.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_picture2.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<%
DIM g_DCa
SET g_DCa = access_dbConnect("Products")



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<title>Migrate DB</title>
</head>

<body>

<%

FUNCTION au_groupIDFromName( oDC, sName )
	au_groupIDFromName = -1
	DIM	sSelect
	sSelect = "" _
	&	"SELECT RID " _
	&	"FROM authorizegroup " _
	&	"WHERE " & dbQ("Name") & " = '" & sName & "' ;"

	DIM	oRS
	SET oRS = dbQueryRead( oDC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		au_groupIDFromName = recNumber( oRS.Fields("RID") )
		oRS.Close
		SET oRS = Nothing
	END IF
END FUNCTION

FUNCTION categoryIDFromName( oDC, sName )
	categoryIDFromName = -1
	DIM	sSelect
	sSelect = "" _
	&	"SELECT CategoryID " _
	&	"FROM categories " _
	&	"WHERE " & dbQ("Name") & " = '" & sName & "' ;"

	DIM	oRS
	SET oRS = dbQueryRead( oDC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		categoryIDFromName = recNumber( oRS.Fields("CategoryID") )
		oRS.Close
		SET oRS = Nothing
	END IF
END FUNCTION

FUNCTION listIDFromName( oDC, sName )
	listIDFromName = -1
	DIM	sNameX
	sNameX = REPLACE(sName,"ending","")
	sNameX = REPLACE(sNameX,"[","(")
	sNameX = REPLACE(sNameX, "]",")")
	DIM	sSelect
	sSelect = "" _
	&	"SELECT RID " _
	&	"FROM pagelistoptions " _
	&	"WHERE " & dbQ("ListName") & " = '" & sNameX & "' ;"

	Response.Write "listIDFromName = " & Server.HTMLEncode(sName) & " = " & Server.HTMLEncode(sNameX) &  "<Br>" & vbCRLF
	Response.Flush

	DIM	oRS
	SET oRS = dbQueryRead( oDC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		listIDFromName = recNumber( oRS.Fields("RID") )
		oRS.Close
		SET oRS = Nothing
	END IF
END FUNCTION

SUB xferAuthorizeUsers()
	DIM	sSelect
	sSelect = "" _
	&	"SELECT * " _
	&	"FROM AuthorizeUsers " _
	&	"WHERE 0 = Disabled; "

	DIM	oRSa
	SET oRSa = access_dbQueryRead( g_DCa, sSelect )
	DIM	oRSax
	IF NOT oRSa IS Nothing THEN
		DIM	oRID
		DIM	oUsername
		DIM	oPassword
		DIM	oEmail
		DIM	oDescription
		DIM	oDateCreated
		DIM	oDateModified
		DIM	oLastname
		DIM	oFirstname
		
		SET oRID = oRSa.Fields("RID")
		SET	oUsername = oRSa.Fields("Username")
		SET oPassword = oRSa.Fields("Password")
		SET oEmail = oRSa.Fields("Email")
		SET	oDescription = oRSa.Fields("Description")
		SET oDateCreated = oRSa.Fields("DateCreated")
		SET oDateModified = oRSa.Fields("DateModified")
		SET oLastname = oRSa.Fields("Lastname")
		SET oFirstname = oRSa.Fields("Firstname")

		DIM	xSelect
		DIM	oRS
		DIM	nRID, n
		DIM	sUsername

		oRSa.MoveFirst
		DO UNTIL oRSa.EOF
			sUsername = access_recString( oUsername )

			xSelect = "" _
				&	"SELECT * " _
				&	"FROM authorizeusers " _
				&	"WHERE Username = '" & sUsername & "' " _
				&	" OR Email LIKE '" & access_recString( oEmail ) & "' ;"

				SET oRS = dbQueryReadEx( g_DC, xSelect )
				IF oRS IS NOTHING THEN
					xSelect = "" _
						&	"SELECT * " _
						&	"FROM authorizeusers " _
						&	"WHERE RID = 0;"
					'oRS.Close
					'SET oRS = Nothing

					SET oRS = dbQueryUpdate( g_DC, xSelect )
					IF NOT oRS IS Nothing THEN
						Response.Write "User = " & sUsername & "<br>" & vbCRLF
						oRS.AddNew
						oRS.Fields("Username").Value = fieldString(sUsername)
						oRS.Fields("Firstname").Value = fieldString(access_recString( oFirstname ))
						oRS.Fields("Lastname").Value = fieldString(access_recString( oLastname ))
						oRS.Fields("Password").Value = fieldString(access_recString( oPassword ))
						oRS.Fields("Email").Value = fieldString(access_recString( oEmail ))
						oRS.Fields("Description").Value = fieldString(access_recString( oDescription ))
						oRS.Fields("DateCreated").Value = fieldDate(access_recDate( oDateCreated ))
						oRS.Fields("DateModified").Value = fieldDate( NOW )
						oRS.Fields("Disabled") = 0
						oRS.Update
						nRID = recNumber( oRS.Fields("RID") )
						oRS.Close
						SET oRS = Nothing


						sSelect = "" _
							&	"SELECT AuthorizeGroup.Name " _
							&	"FROM AuthorizeGroup " _
							&	"INNER JOIN AuthorizeGroupMap ON AuthorizeGroupMap.GroupID = AuthorizeGroup.RID " _
							&	"WHERE AuthorizeGroupMap.UserID=" & access_recNumber( oRID ) & ";"
						xSelect = "" _
							&	"SELECT * " _
							&	"FROM authorizegroupmap " _
							&	"WHERE RID = 0 ;"
						Response.Write "---Connecting Groups<br>"
						SET oRSax = access_dbQueryRead( g_DCa, sSelect )
						IF NOT oRSax IS Nothing THEN
							oRSax.MoveFirst
							DO UNTIL oRSax.EOF
								' create new relations in mysql
								Response.Write "---Group=" & access_recString( oRSax.Fields("Name") ) & "<br>"
								n = au_groupIDFromName( g_DC, access_recString( oRSax.Fields("Name") ) )
								IF 0 < n THEN
									SET oRS = dbQueryUpdate( g_DC, xSelect )
									IF NOT oRS IS Nothing THEN
										oRS.AddNew
										oRS.Fields("UserID").Value = nRID
										oRS.Fields("GroupID").Value = n
										oRS.Update
										oRS.Close
										SET oRS = Nothing
									END IF
								END IF
									
								oRSax.MoveNext
							LOOP
							oRSax.Close
							SET oRSax = Nothing
						END IF

					END IF
					
				ELSE
					oRS.Close
					SET oRS = Nothing
					Response.Write "User already migrated = " & access_recString( oUsername ) & "<br>" & vbCRLF
				END IF
				
			oRSa.MoveNext
		LOOP
	ELSE
		Response.Write "Unable to access user authorization data<br>" & vbCRLF
	END IF
	Response.Flush
END SUB


SUB xferCategories
	DIM	sSelect
	sSelect = "" _
	&	"SELECT * " _
	&	"FROM Categories ; "

	DIM	xSelect
	xSelect = "" _
	&	"SELECT * " _
	&	"FROM categories " _
	&	"WHERE 0 = CategoryID; "

	DIM	sNameSelect

	DIM	oRSa
	DIM	oRS
	SET oRSa = access_dbQueryRead( g_DCa, sSelect )
	IF NOT oRSa IS Nothing THEN
		oRSa.MoveFirst
		DO UNTIL oRSa.EOF
			sNameSelect = "" _
			&	"SELECT CategoryID " _
			&	"FROM categories " _
			&	"WHERE Name = '" & access_recString(oRSa.Fields("Name")) & "';"
			SET oRS = dbQueryReadEx( g_DC, sNameSelect )
			IF oRS IS Nothing THEN
				'oRS.Close
				'SET oRS = Nothing
				SET oRS = dbQueryUpdate( g_DC, xSelect )
				IF NOT oRS IS Nothing THEN
					Response.Write "Category=" & access_recString(oRSa.Fields("Name")) & "<br>" & vbCRLF
					oRS.AddNew
					oRS.Fields("Name").Value = fieldString(access_recString(oRSa.Fields("Name")))
					oRS.Fields("SortName").Value = fieldString(access_recString(oRSa.Fields("SortName")))
					oRS.Fields("ShortName").Value = fieldString(access_recString(oRSa.Fields("ShortName")))
					oRS.Fields("ShortDesc").Value = fieldString(access_recString(oRSa.Fields("ShortDesc")))
					oRS.Fields("Disabled").Value = fieldBool(access_recBool(oRSa.Fields("Disabled")))
					oRS.Update
					oRS.Close
					SET oRS = Nothing
				END IF
			ELSE
				Response.Write "Category already migrated = "  & access_recString(oRSa.Fields("Name")) & "<br>" & vbCRLF
			END IF
			
			oRSa.MoveNext
		LOOP
		oRSa.Close
		SET oRSa = Nothing
	ELSE
		Response.Write "Unable to access category data<br>" & vbCRLF
	END IF
	Response.Flush
END SUB


SUB xferTabs
	DIM	sSelect
	DIM	xSelect
	DIM	sNameSelect
	DIM	nTabID
	DIM	n

	DIM	oRSa
	DIM	oRSax
	DIM	oRS

	' first clear any predefined tabs so that only user tabs are created
	sSelect = "" _
	&	"SELECT * " _
	&	"FROM tabs " _
	&	";"
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		DO UNTIL oRS.EOF
			oRS.Delete adAffectCurrent
			oRS.MoveNext
		LOOP
		oRS.Close
		SET oRS = Nothing
	END IF

	sSelect = "" _
	&	"SELECT * " _
	&	"FROM Tabs; "
	SET oRSa = access_dbQueryRead( g_DCa, sSelect )
	IF NOT oRSa IS Nothing THEN
		oRSa.MoveFirst
		DO UNTIL oRSa.EOF
			sNameSelect = "" _
			&	"SELECT TabID " _
			&	"FROM tabs " _
			&	"WHERE Name = '" & access_recString(oRSa.Fields("Name")) & "';"
			SET oRS = dbQueryReadEx( g_DC, sNameSelect )
			IF oRS IS Nothing THEN
				'oRS.Close
				'SET oRS = Nothing
				xSelect = "" _
				&	"SELECT * " _
				&	"FROM tabs " _
				&	"WHERE 0 = TabID; "
				SET oRS = dbQueryUpdate( g_DC, xSelect )
				IF NOT oRS IS Nothing THEN
					Response.Write "Tab=" & access_recString(oRSa.Fields("Name")) & "<br>" & vbCRLF
					oRS.AddNew
					oRS.Fields("Name").Value = fieldString(access_recString(oRSa.Fields("Name")))
					oRS.Fields("SortName").Value = fieldString(access_recString(oRSa.Fields("SortName")))
					oRS.Fields("Description").Value = fieldString(access_recString(oRSa.Fields("Description")))
					oRS.Fields("Disabled").Value = fieldBool(access_recBool(oRSa.Fields("Disabled")))
					oRS.Update
					nTabID = recNumber(oRS.Fields("TabID"))
					oRS.Close
					SET oRS = Nothing


					sSelect = "" _
						&	"SELECT Categories.Name " _
						&	"FROM Categories " _
						&	"INNER JOIN TabCategoryMap ON TabCategoryMap.CategoryID = Categories.CategoryID " _
						&	"WHERE TabCategoryMap.TabID=" & access_recNumber( oRSa.Fields("TabID") ) & ";"
					xSelect = "" _
						&	"SELECT * " _
						&	"FROM tabcategorymap " _
						&	"WHERE RID = 0 ;"
					Response.Write "---Connecting Categories<br>"
					SET oRSax = access_dbQueryRead( g_DCa, sSelect )
					IF NOT oRSax IS Nothing THEN
						oRSax.MoveFirst
						DO UNTIL oRSax.EOF
							' create new relations in mysql
							Response.Write "---Category=" & access_recString( oRSax.Fields("Name") ) & "<br>"
							n = categoryIDFromName( g_DC, access_recString( oRSax.Fields("Name") ) )
							IF 0 < n THEN
								SET oRS = dbQueryUpdate( g_DC, xSelect )
								IF NOT oRS IS Nothing THEN
									oRS.AddNew
									oRS.Fields("TabID").Value = nTabID
									oRS.Fields("CategoryID").Value = n
									oRS.Update
									oRS.Close
									SET oRS = Nothing
								END IF
							END IF
									
							oRSax.MoveNext
						LOOP
						oRSax.Close
						SET oRSax = Nothing
					END IF


					sSelect = "" _
						&	"SELECT PageListOptions.ListName " _
						&	"FROM PageListOptions " _
						&		"INNER JOIN PageListMap ON PageListMap.ListID = PageListOptions.RID " _
						&	"WHERE PageListMap.TabID = " & access_recNumber( oRSa.Fields("TabID") ) _
						&	";"
					xSelect = "" _
						&	"SELECT * " _
						&	"FROM pagelistmap " _
						&	"WHERE RID = 0 ;"
					SET oRSax = access_dbQueryRead( g_DCa, sSelect )
					IF NOT oRSax IS Nothing THEN
						n = listIDFromName( g_DC, access_recString( oRSax.Fields("ListName") ) )
						oRSax.Close
						SET oRSax = Nothing
						IF 0 < n THEN
							SET oRS = dbQueryUpdate( g_DC, xSelect )
							IF NOT oRS IS Nothing THEN
								oRS.AddNew
								oRS.Fields("TabID").Value = nTabID
								oRS.Fields("ListID").Value = n
								oRS.Update
								oRS.Close
								SET oRS = Nothing
							END IF
						END IF
					END IF


				END IF
			ELSE
				Response.Write "Tab already migrated = "  & access_recString(oRSa.Fields("Name")) & "<br>" & vbCRLF
			END IF
			
			oRSa.MoveNext
		LOOP
		oRSa.Close
		SET oRSa = Nothing
	ELSE
		Response.Write "Unable to access tabs data<br>" & vbCRLF
	END IF
	Response.Flush
END SUB


SUB xferPages
	DIM	sSelect
	DIM	xSelect
	DIM	sNameSelect
	DIM	nPageID
	DIM	n

	DIM	oRSa
	DIM	oRSax
	DIM	oRS

	DIM	sPicture
	DIM	sPictFile

	sSelect = "" _
	&	"SELECT Pages.* " _
	&	", Categories.Name AS CatName " _
	&	", Schedules.DateBegin " _
	&	", Schedules.DateEnd " _
	&	", Schedules.DateEvent " _
	&	", Schedules.Disabled AS SchedDisabled " _
	&	"FROM (Pages " _
	&	"INNER JOIN Categories ON Categories.CategoryID = Pages.Category) " _
	&	"LEFT JOIN Schedules ON Schedules.PageID = Pages.RID " _
	&	"; "
	SET oRSa = access_dbQueryRead( g_DCa, sSelect )
	IF NOT oRSa IS Nothing THEN
		oRSa.MoveFirst
		DO UNTIL oRSa.EOF
			sNameSelect = "" _
			&	"SELECT RID " _
			&	"FROM pages " _
			&	"WHERE PageName = '" & access_recString(oRSa.Fields("PageName")) & "';"
			SET oRS = dbQueryReadEx( g_DC, sNameSelect )
			IF oRS IS Nothing THEN
				'oRS.Close
				'SET oRS = Nothing
				xSelect = "" _
				&	"SELECT * " _
				&	"FROM pages " _
				&	"WHERE 0 = RID; "
				SET oRS = dbQueryUpdate( g_DC, xSelect )
				IF NOT oRS IS Nothing THEN
					Response.Write "Page=" & access_recString(oRSa.Fields("PageName")) & "<br>" & vbCRLF
					oRS.AddNew
					oRS.Fields("PageName").Value = fieldString(access_recString(oRSa.Fields("PageName")))
					oRS.Fields("SortName").Value = fieldString(access_recString(oRSa.Fields("SortName")))
					oRS.Fields("Format").Value = fieldString(access_recString(oRSa.Fields("Format")))
					oRS.Fields("Title").Value = fieldString(access_recString(oRSa.Fields("Title")))
					oRS.Fields("Description").Value = fieldString(access_recString(oRSa.Fields("Description")))
					oRS.Fields("Body").Value = fieldString(access_recString(oRSa.Fields("Body")))
					oRS.Fields("DateCreated").Value = fieldDate(access_recDate(oRSa.Fields("DateCreated")))
					oRS.Fields("DateModified").Value = fieldDate(access_recDate(oRSa.Fields("DateModified")))
					oRS.Fields("Category").Value = categoryIDFromName(g_DC, access_recString(oRSa.Fields("CatName")))
					sPicture = access_recString(oRSa.Fields("Picture"))
					IF 0 < LEN(sPicture) THEN
						sPictFile = picture_buildPath( "Pages", sPicture )
						IF g_oFSO.FileExists(sPictFile) THEN
							n = storeImage( g_DC, sPictFile )
							IF 0 < n THEN
								oRS.Fields("Picture").Value = n
							END IF
						END IF
					END IF
					oRS.Fields("Disabled").Value = fieldBool(access_recBool(oRSa.Fields("Disabled")))
					oRS.Update
					nPageID = recNumber(oRS.Fields("RID"))
					oRS.Close
					SET oRS = Nothing

					IF (NOT ISNULL(oRSa.Fields("DateBegin"))) _
						OR (NOT ISNULL(oRSa.Fields("DateEnd")) ) _
						OR (NOT ISNULL(oRSa.Fields("DateEvent")) ) _
						THEN

						xSelect = "" _
						&	"SELECT * " _
						&	"FROM schedules " _
						&	"WHERE 0=RID; "
						SET oRS = dbQueryUpdate( g_DC, xSelect )
						IF NOT oRS IS Nothing THEN
							oRS.AddNew
							oRS.Fields("PageID").Value = nPageID
							oRS.Fields("DateBegin").Value = fieldDate(access_recDate(oRSa.Fields("DateBegin")))
							oRS.Fields("DateEnd").Value = fieldDate(access_recDate(oRSa.Fields("DateEnd")))
							oRS.Fields("DateEvent").Value = fieldDate(access_recDate(oRSa.Fields("DateEvent")))
							oRS.Update
							oRS.Close
							SET oRS = Nothing
						END IF
					END IF

					xSelect = "" _
					&	"SELECT * " _
					&	"FROM Pictures " _
					&	"WHERE PageID=" & access_recNumber( oRSA.Fields("RID") ) _
					&	";"
					SET oRSax = access_dbQueryRead( g_DCa, xSelect )
					IF NOT oRSax IS Nothing THEN
						xSelect = "" _
						&	"SELECT * " _
						&	"FROM pictures " _
						&	"WHERE 0=PictureID " _
						&	";"
						oRSax.MoveFirst
						DO UNTIL oRSax.EOF
							sPicture = access_recString(oRSax.Fields("File"))
							IF 0 < LEN(sPicture) THEN
								sPictFile = picture_buildPath( "Pages", sPicture )
								IF g_oFSO.FileExists(sPictFile) THEN
									n = storeImage( g_DC, sPictFile )
									IF 0 < n THEN
										SET oRS = dbQueryUpdate( g_DC, xSelect )
										IF NOT oRS IS Nothing THEN
											oRS.AddNew
											oRS.Fields("Label").Value = fieldString( access_recString(oRSax.Fields("Label")) )
											oRS.Fields("PageID").Value = nPageID
											oRS.Fields("ImageID").Value = n
											oRS.Update
											oRS.Close
											SET oRS = Nothing
										END IF
									END IF
								END IF
							END IF
							oRSax.MoveNext
						LOOP
						oRSax.Close
						SET oRSax = Nothing
					END IF

				END IF
			ELSE
				Response.Write "Page already migrated = "  & access_recString(oRSa.Fields("PageName")) & "<br>" & vbCRLF
			END IF
			
			oRSa.MoveNext
		LOOP
		oRSa.Close
		SET oRSa = Nothing
	ELSE
		Response.Write "Unable to access page data<br>" & vbCRLF
	END IF
	Response.Flush

END SUB


SUB xferLibraryPictures
	DIM	sSelect
	DIM	xSelect
	DIM	sNameSelect
	DIM	nPageID
	DIM	n

	DIM	oRSa
	DIM	oRSax
	DIM	oRS

	DIM	sPicture
	DIM	sPictFile

	sSelect = "" _
	&	"SELECT RID " _
	&	"FROM pages " _
	&	"INNER JOIN categories ON categories.CategoryID = pages.Category " _
	&	"WHERE categories.Name = '@Library' AND pages.PageName = '~Library' " _
	&	";"
	SET oRS = dbQueryRead( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		oRS.MoveFirst
		nPageID = recNumber( oRS.Fields("RID") )

		xSelect = "" _
		&	"SELECT * " _
		&	"FROM Pictures " _
		&	"WHERE PageID=0"  _
		&	";"
		SET oRSax = access_dbQueryRead( g_DCa, xSelect )
		IF NOT oRSax IS Nothing THEN
			xSelect = "" _
			&	"SELECT * " _
			&	"FROM pictures " _
			&	"WHERE 0=PictureID " _
			&	";"
			oRSax.MoveFirst
			DO UNTIL oRSax.EOF
				sPicture = access_recString(oRSax.Fields("File"))
				IF 0 < LEN(sPicture) THEN
					sPictFile = picture_buildPath( "Pages", sPicture )
					IF g_oFSO.FileExists(sPictFile) THEN
						n = storeImage( g_DC, sPictFile )
						IF 0 < n THEN
							SET oRS = dbQueryUpdate( g_DC, xSelect )
							IF NOT oRS IS Nothing THEN
								oRS.AddNew
								oRS.Fields("Label").Value = fieldString( access_recString(oRSax.Fields("Label")) )
								oRS.Fields("PageID").Value = nPageID
								oRS.Fields("ImageID").Value = n
								oRS.Update
								oRS.Close
								SET oRS = Nothing
							END IF
						END IF
					END IF
				END IF
				oRSax.MoveNext
			LOOP
			oRSax.Close
			SET oRSax = Nothing
		END IF

	ELSE
		Response.Write "Unable to access library picture data<br>" & vbCRLF
	END IF
	Response.Flush

END SUB

SUB	updateConfig( oDC, sName, sValue )
	DIM	sSelect
	DIM	s
	s = sValue
	sSelect = "" _
	&	"SELECT * " _
	&	"FROM config " _
	&	"WHERE Name = '" & sName & "' " _
	&	";"
	DIM	oRS
	SET oRS = dbQueryUpdate( oDC, sSelect )
	IF NOT oRS IS Nothing THEN
		IF ISNULL(s) THEN
			s = ""
		END IF
		IF CSTR(s) <> recString(oRS.Fields("DataValue")) THEN
			oRS.Fields("DataValue").Value = fieldString(CSTR(s))
			oRS.Update
		END IF
		oRS.Close
		SET oRS = Nothing
	END IF
END SUB


SUB xferConfig
	DIM	sSelect
	DIM	oRSa

	sSelect = "" _
	&	"SELECT * " _
	&	"FROM Config " _
	&	";"
	SET oRSa = access_dbQueryRead( g_DCa, sSelect )
	IF NOT oRSa IS Nothing THEN
		updateConfig g_DC, "Domain", oRSa.Fields("Domain").Value
		updateConfig g_DC, "SiteName", oRSa.Fields("SiteName").Value
		updateConfig g_DC, "SiteTitle", oRSa.Fields("SiteTitle").Value
		updateConfig g_DC, "ShortSiteName", oRSa.Fields("SiteOrgDesignation").Value
		updateConfig g_DC, "SiteMotto", oRSa.Fields("SiteMotto").Value
		updateConfig g_DC, "SiteCity", oRSa.Fields("SiteLocation").Value
		updateConfig g_DC, "CopyrightStartYear", oRSa.Fields("CopyrightStartYear").Value
		updateConfig g_DC, "MailServer", oRSa.Fields("MailServer").Value
		updateConfig g_DC, "MailboxUpcomingEvents", oRSa.Fields("MailboxUpcomingEvents").Value
		updateConfig g_DC, "MailboxNewsletter", oRSa.Fields("MailboxNewsletter").Value
		updateConfig g_DC, "AnnonUser", oRSa.Fields("MailUser").Value
		updateConfig g_DC, "AnnonPW", oRSa.Fields("MailPW").Value
		updateConfig g_DC, "SiteZip", oRSa.Fields("SiteZip").Value
		updateConfig g_DC, "SiteChapter", oRSa.Fields("SiteChapter").Value
		updateConfig g_DC, "SiteChapterID", oRSa.Fields("SiteChapterID").Value
		updateConfig g_DC, "SiteTabNewsletters", oRSa.Fields("SiteTabNewsletters").Value
		updateConfig g_DC, "SiteTabAnnouncements", oRSa.Fields("SiteTabAnnouncements").Value
		updateConfig g_DC, "SiteRSSAnnounceInclude", oRSa.Fields("SiteRSSAnnounceInclude").Value
		updateConfig g_DC, "SiteRSSAnnounceExclude", oRSa.Fields("SiteRSSAnnounceExclude").Value
		oRSa.Close
		SET oRSa = Nothing
	ELSE
		Response.Write "Unable to access configuration data<br>" & vbCRLF
	END IF
	Response.Flush
END SUB


IF NOT g_DCa IS Nothing THEN
	IF NOT g_DC IS Nothing THEN
		xferAuthorizeUsers
		xferCategories
		xferTabs
		xferPages
		xferLibraryPictures
		xferConfig

		cache_clearFolders
	ELSE
		Response.Write "Target database is not opened<br>"
	END IF
ELSE
	Response.Write "MS-Access database is not open<br>"
END IF


%>

</body>

</html>
<%
SET g_DCa = Nothing
%>
<!--#include file="scripts\include_db_end.asp"-->
