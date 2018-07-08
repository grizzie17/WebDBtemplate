<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM g_oFSO
SET g_oFSO = CreateObject( "Scripting.FileSystemObject" )



%>
<!--#include file="scripts\config.asp"-->
<!--#include file="config\configuser.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts\include_announce.asp"-->
<!--#include file="scripts/index_tools.asp"-->
<%


Response.ContentType = "text/xml"

Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCRLF




FUNCTION formatUTC( d )

	DIM	s
	DIM	t
	
	'Tue, 27 Mar 2007 08:56:19 CST
	
	s = WEEKDAYNAME( WEEKDAY(d), TRUE )
	s = s & ", "
	s = s & DATEPART( "d", d )
	s = s & " "
	s = s & MONTHNAME( MONTH(d), TRUE )
	s = s & " "
	s = s & DATEPART( "yyyy", d )
	s = s & " "
	t = TRIM(LEFT(FORMATDATETIME( d, vbLongTime ), 8))
	s = s & t
	s = s & " UTC"
	
	formatUTC = s

END FUNCTION




SUB makeChannelHeader()

	Response.Write "<title>" & Server.HTMLEncode(g_sSiteName) & "</title>" & vbCRLF
	
	DIM	i
	DIM	sURL
	sURL = Request.ServerVariables("URL")
	i = INSTRREV(sURL,"/")
	IF ( 0 < i ) THEN sURL = LEFT(sURL,i)
	sURL = "http://" & Request.ServerVariables("HTTP_HOST") & sURL
	
	Response.Write "<link>" & sURL & "</link>" & vbCRLF
	
	Response.Write "<description>Announcements</description>" & vbCRLF
	Response.Write "<language>en-us</language>" & vbCRLF



END SUB


SUB processItems()


	DIM	i
	DIM	sURL
	sURL = Request.ServerVariables("URL")
	i = INSTRREV(sURL,"/")
	IF ( 0 < i ) THEN sURL = LEFT(sURL,i)
	sURL = "http://" & Request.ServerVariables("HTTP_HOST") & sURL



	g_sUseFileNameSuffix = ".htm"
	buildFileList "announcements", TRUE
	appendFileList 0, -21

	IF 0 < nFileCount THEN
	
	
		'aFileSplit = SPLIT( aFileList(0), vbTAB )
		'dAnnouncementsModified = CDATE( aFileSplit(kFI_DateLastModified) )
	

		DIM	sJPGFile
		DIM	oFile
		DIM	sLink
		DIM	sMime
		DIM	sSrcUpdate
		DIM	sAltUpdate
		FOR nLen = 0 TO nFileCount-1
			aFileSplit = SPLIT( aFileList(nLen), vbTAB, -1, vbTextCompare )
			
			sLink = sURL & "announcements.asp?page=" & aFileSplit(kFI_Name)
			
			Response.Write "<item>" &vbCRLF
			Response.Write "<title>"
			Response.Write Server.HTMLEncode(aFileSplit(kFI_Title)) 
			Response.Write "</title>" & vbCRLF
			Response.Write "<link>" & sLink & "</link>" & vbCRLF
			Response.Write "<guid isPermaLink=""true"">" & sLink & "</guid>" & vbCRLF
			IF 0 < LEN(aFileSplit(kFI_Description)) THEN
				Response.Write "<description>"
				Response.Write Server.HTMLEncode(aFileSplit(kFI_Description))
				Response.Write "</description>" & vbCRLF
			END IF
			IF 0 < LEN(aFileSplit(kFI_Keywords)) THEN
				Response.Write "<category>"
				Response.Write Server.HTMLEncode(aFileSplit(kFI_Keywords))
				Response.Write "</category>" & vbCRLF
			END IF


'createdDate -- if present as a sub-element of <item>, it specifies the date that the item content was created. Multiple createdDate elements are not allowed.

'modifiedDate -- if present as a sub-element of <item>, it specifies the date that the item content was last updated. Multiple modifiedDate elements are not allowed.

'startDate -- if present as a sub-element of <item>, it specifies the start date of the event item. Multiple startDate elements are not allowed.

'endDate -- if present as a sub-element of <item>, it specifies the end date of the event item. Multiple endDate elements are not allowed.

			'IF 0 < LEN(aFileSplit(kFI_SortName)) THEN
			'	Response.Write "<xBCG:sortName>" & aFileSplit(kFI_SortName) & "</xBCG:sortName>" & vbCRLF
			'END IF
			'IF 0 < LEN(aFileSplit(kFI_DateRange)) THEN
			'	Response.Write "<xBCG:dateRange>" & aFileSplit(kFI_DateRange) & "</xBCG:dateRange>" & vbCRLF
			'END IF
			IF 0 < LEN(aFileSplit(kFI_EventDate)) THEN
				Response.Write "<event:startDate>" & formatUTC( CDATE(aFileSplit(kFI_EventDate))) & "</event:startDate>" & vbCRLF
				Response.Write "<pubDate>"
				Response.Write formatUTC( CDATE(aFileSplit(kFI_EventDate)))
				Response.Write "</pubDate>" & vbCRLF
			ELSE
				Response.Write "<pubDate>"
				Response.Write formatUTC( CDATE( aFileSplit(kFI_DateLastModified) ))
				Response.Write "</pubDate>" & vbCRLF
			END IF

			sJPGFile = ""
			sJPGFile = REPLACE( aFileSplit(kFI_Name), ".htm", ".jpg" )
			IF NOT g_oFSO.FileExists( Server.MapPath( "announcements/" & sJPGFile ) ) THEN
				sJPGFile = REPLACE( aFileSplit(kFI_Name), ".htm", ".gif" )
				IF NOT g_oFSO.FileExists( Server.MapPath( "announcements/" & sJPGFile ) ) THEN
					sJPGFile = REPLACE( aFileSplit(kFI_Name), ".htm", ".png" )
					IF NOT g_oFSO.FileExists( Server.MapPath( "announcements/" & sJPGFile ) ) THEN
						sJPGFile = ""
					ELSE
						sMime = "image/png"
					END IF
				ELSE
					sMime = "image/gif"
				END IF
			ELSE
				sMime = "image/jpg"
			END IF
			IF "" <> sJPGFile THEN
				SET oFile = g_oFSO.GetFile( Server.MapPath( "announcements/" & sJPGFile ))
				Response.Write "<enclosure "
				Response.Write "url=""" & sURL & "announcements/" & sJPGFile & """ "
				Response.Write "length=""" & oFile.size & """ "
				Response.Write "type=""" & Server.HTMLEncode(sMime) & """ />" &vbCRLF
				SET oFile = Nothing
			END IF

			Response.Write "</item>" &vbCRLF
		NEXT 'nLen
	END IF

END SUB



SUB makeChannel()
	Response.Write "<channel>" & vbCRLF
	
	
	
	makeChannelHeader()
	
	processItems()
	
	
	
	Response.Write "</channel>" & vbCRLF

END SUB


SUB makeRSS()

	Response.Write "<rss version=""2.0""  xmlns:event=""http://gsc-webserver.mit.edu/eventRssModule"">" & vbCRLF
	
	
	makeChannel()
	
	
	Response.Write "</rss>" & vbCRLF


END SUB

makeRSS



SET g_oFSO = Nothing

%>