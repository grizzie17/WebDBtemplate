<%


DIM	g_emailnotify_processHtmlPicture
g_emailnotify_processHtmlPicture = "notify_f_processHtmlPicture"




FUNCTION notify_f_processHtmlPicture( oOut, sFilename )

	notify_f_processHtmlPicture = "" _
			&	"<img border=""0"" src=""" & virtualFromPhysicalPath(sFilename) & """>"

END FUNCTION

FUNCTION notifySMTP_f_processHtmlPicture( oOut, sFilename )

	DIM sPictID
	sPictID = oOut.AddInlineAttachment( sFilename )
	IF "" <> sPictID THEN
		notifySMTP_f_processHtmlPicture = "" _
				&	"<img border=""0"" src=""cid:" & sPictID & """>"
	ELSE
		notifySMTP_f_processHtmlPicture = ""
	END IF

END FUNCTION







FUNCTION formatPostDate( sDate )

	DIM	sD
	
	sD = MID(sDate,5,2) & "/" & MID(sDate,7,2) & "/" & MID(sDate,1,4)
	
	DIM	sSfx
	sSfx = "am"
	DIM	nH
	nH = CINT( MID(sDate,9,2) )
	IF nH < 1 THEN
		nH = 12
	ELSEIF 11 < nH THEN
		sSfx = "pm"
		IF 12 < nH THEN nH = nH - 12
	END IF
	
	sD = sD & " " & nH & ":" & MID(sDate,11,2) & ":" & MID(sDate,13,2) & " " & sSfx
	
	DIM	d
	d = CDATE( sD )
	formatPostDate = formatUTC( d )

END FUNCTION




FUNCTION formatUTC( d )

	DIM	s
	DIM	t
	
	'Tue, 27 Mar 2007 08:56:19 CST
	
	t = TRIM(LEFT(FORMATDATETIME( d, vbLongTime ), 8))
	s = "" _
		&	WEEKDAYNAME( WEEKDAY(d), TRUE ) _
		&	", " _
		&	DATEPART( "d", d ) _
		&	" " _
		&	MONTHNAME( MONTH(d), TRUE ) _
		&	" " _
		&	DATEPART( "yyyy", d ) _
		&	" " _
		&	t _
		&	" UTC"
	
	formatUTC = s

END FUNCTION





DIM g_aWD
g_aWD = ARRAY( "su", "mo", "tu", "we", "th", "fr", "sa" )


FUNCTION lookupWD( s )

	lookupWD = -1
	DIM	i
	FOR i = LBOUND(g_aWD) TO UBOUND(g_aWD)
		IF s = g_aWD(i) THEN
			lookupWD = i + 1
			EXIT FOR
		END IF
	NEXT 'i

END FUNCTION

FUNCTION getTargetWD( aXD(), nDay )

	getTargetWD = aXD(UBOUND(aXD))
	DIM	i
	FOR i = UBOUND(aXD) to LBOUND(aXD) STEP -1
		IF nDay < aXD(i) THEN
			getTargetWD = aXD(i)
			EXIT FOR
		ELSEIF nDay = aXD(i) THEN
			IF i < UBOUND(aXD) THEN
				getTargetWD = aXD(i+1)
				EXIT FOR
			ELSE
				getTargetWD = aXD(LBOUND(aXD))
				EXIT FOR
			END IF
		END IF
	NEXT 'i

END FUNCTION


SUB getNotificationDateValues( nDay, d )

	DIM	aWD
	aWD = SPLIT( g_sUpcomingEventsSchedule, "," )
	
	DIM aXD()
	REDIM aXD(UBOUND(aWD))
	
	DIM	i
	
	FOR i = LBOUND(aWD) TO UBOUND(aWD)
		aXD(i) = lookupWD( aWD(i) )
	NEXT 'i

	'DIM	d
	d = DATEADD("h", g_nServerTimeZoneOffset, NOW)
	
	'DIM	nDay
	DIM	nTargetDay
	nDay = WEEKDAY( d )
	nTargetDay = getTargetWD( aXD, nDay )
	
	nDay = (nTargetDay - nDay + 7) MOD 7

END SUB


FUNCTION findNearestNotificationDate()

	DIM	nDay, d
	getNotificationDateValues nDay, d

	findNearestNotificationDate = DATEADD( "d", nDay, d )

END FUNCTION


FUNCTION findNextNotificationDate()

	DIM	nDay, d
	getNotificationDateValues nDay, d
	IF 0 = nDay THEN nDay = 7
	
	findNextNotificationDate = DATEADD( "d", nDay, d )
		
END FUNCTION










DIM g_sProcessAnnouncementsPictureFunc
g_sProcessAnnouncementsPictureFunc = "f_processNotifyAnnouncementsPicture"


FUNCTION f_processNotifyAnnouncementsPicture( sPicture )
	f_processNotifyAnnouncementsPicture = "picture.asp?table=pages&name=" & sPicture
END FUNCTION


FUNCTION processNotifyAnnouncementsPicture( sPicture )
	processNotifyAnnouncementsPicture = EVAL( g_sProcessAnnouncementsPictureFunc & "( sPicture )")
END FUNCTION


FUNCTION processNotifyAnnouncements( oOut, dToday )

	processNotifyAnnouncements = ""

	DIM	nID
	nID = defineTabSortDetails( g_sHomepageAnnouncements )
	
	DIM	oRS
	SET oRS = getPageListRS( nID )

	IF NOT oRS IS Nothing THEN
		IF 0 < oRS.RecordCount THEN
		
			DIM	oPageRID
			DIM	oFormat
			DIM	oTitle
			DIM	oDescription
			DIM	oPicture
			DIM	oDateModified
			DIM	oCategory
			DIM	oTabID
			DIM	oDateBegin
			DIM	oDateEnd
			DIM	oDateEvent
			
			SET oPageRID = oRS.Fields("RID")
			SET oFormat = oRS.Fields("Format")
			SET oTitle = oRS.Fields("Title")
			SET oDescription = oRS.Fields("Description")
			SET oPicture = oRS.Fields("Picture")
			SET oDateModified = oRS.Fields("DateModified")
			SET oCategory = oRS.Fields("Category")
			SET oTabID = oRS.Fields("TabID")
			SET oDateBegin = oRS.Fields("DateBegin")
			SET oDateEnd = oRS.Fields("DateEnd")
			SET oDateEvent = oRS.Fields("DateEvent")
			
			DIM	nPageRID
			DIM	sFormat
			DIM	sTitle
			DIM	sDescription
			DIM	sPicture
			DIM	dDateModified
			DIM	nCategory
			DIM	nTabID
			DIM	dDateBegin
			DIM	dDateEnd
			DIM	dDateEvent
			
			DIM	sSrcUpdate
			DIM	sAltUpdate
			DIM	sHTMLTitle
			
			
			DIM	sOutPicture
			DIM	sOutDescription
			DIM	sOutTitle
			DIM	sOutUpdated
			
			DIM	sURL
			
			DIM	bProcess
			DIM	dEarlyDate
			DIM	dLateDate
			DIM	dTonight
			DIM	f
		
			dTonight = CDATE(MONTH(dToday)&"/"&DAY(dToday)&"/"&YEAR(dToday)&" 11:59:59 PM")
			
			dEarlyDate = DATEADD( "d", -g_UE_nDurationAnnounce, dToday )
			dLateDate = DATEADD( "d", g_UE_nDurationAnnounce, dToday )
			
			'Response.Write "Today = " & dToday & "<br>"
			'Response.Write "Tonight = " & dTonight & "<br>"
			'Response.Write "Early = " & dEarlyDate & "<br>"
			'Response.Write "Late = " & dLateDate & "<br>"
		
			oRS.MoveFirst
			DO UNTIL oRS.EOF
			
				sOutPicture = ""
				sOutDescription = ""
				sOutTitle = ""
				sOutUpdated = ""
				
				bProcess = FALSE
			
				nPageRID = recNumber(oPageRID)
				sFormat = recString(oFormat)
				sTitle = recString(oTitle)
				sDescription = recString(oDescription)
				sPicture = recString(oPicture)
				dDateModified = recDate(oDateModified)
				nCategory = recNumber(oCategory)
				nTabID = recNumber(oTabID)
				dDateBegin = recDate(oDateBegin)
				dDateEnd = recDate(oDateEnd)
				dDateEvent = recDate(oDateEvent)
				
				'Response.Write "Begin = " & dDateBegin & "<br>"
				'Response.Write "End = " & dDateEnd & "<br>"
				'Response.Write "Event = " & dDateEvent & "<br>"
				'Response.Write "Modified = " & dDateModified & "<br>"
				
				IF CDATE( "12:00:00AM" ) <> dDateEvent THEN
					'Response.Write "diff-event = " & DATEDIFF( "n", dEarlyDate, dDateEvent ) & "<br>"
					IF 0 < DATEDIFF( "n", dEarlyDate, dDateEvent ) THEN
						'Response.Write "diff-event2 = " & DATEDIFF( "n", dTonight, dDateEvent ) & "<br>"
						IF 0 < DATEDIFF( "n", dTonight, dDateEvent ) THEN
							bProcess = TRUE
						END IF
					END IF
				END IF
				f = DATEDIFF( "n", dDateModified, dLateDate )
				f = f / (60 * 24)
				'Response.Write "diff-mod = " & f & "<br>"
				IF 0 <= f  AND  f <= g_UE_nDurationAnnounce THEN bProcess = TRUE
				
				
				IF bProcess THEN
				
					sURL = "page.asp?tab=" & nTabID & "&category=" & nCategory & "$page=" & nPageRID
					
					sOutTitle = "" _
						&	"<div class=""announcementstitle"">" & vbCRLF _
						&	"<a href=""" & sURL & """>" _
						&	pagebody_formatString( sFormat, sTitle ) _
						&	"</a>" _
						&	"</div>" & vbCRLF
	
	
	
					IF "" <> sPicture THEN
						sOutPicture = "" _
							&	"<div>" _
							&	"<a href=""" & sURL & """>" _
							&	"<img border=""0"" alt=""" & sHTMLTitle & """ src=""" & processNotifyAnnouncementsPicture(sPicture) & """>" _
							&	"</a>" & vbCRLF _
							&	"</div>" & vbCRLF
					END IF
					IF 0 < LEN(sDescription) THEN
						sOutDescription = "" _
							&	"<div align=""left"">" _
							&	pagebody_formatString( sFormat, sDescription ) _
							&	"&nbsp;&nbsp; <a href=""" & sURL & """>" _
							&	"More..." _
							&	"</a>" & vbCRLF _
							&	"</div>" & vbCRLF
					END IF
					
					sOutUpdated = "" _
						&	"<div style=""text-align: left; color: #999999; font-family: sans-serif; font-size: xx-small;"">" _
						&	"Updated: " & Server.HTMLEncode(DATEADD("h", g_nServerTimeZoneOffset, dDateModified)) _
						&	"</div><br>" & vbCRLF
	
					processNotifyAnnouncements = processNotifyAnnouncements _
						&	sOutTitle _
						&	sOutPicture _
						&	sOutDescription _
						&	sOutUpdated
	
				END IF

				oRS.MoveNext
			LOOP
			
			
			IF "" <> processNotifyAnnouncements THEN
				processNotifyAnnouncements = "" _
					&	"<hr>" & vbCRLF _
					&	"<h2 class=""emailnotification"">New or Updated Announcements/Articles</h2>" & vbCRLF _
					&	processNotifyAnnouncements & vbCRLF _
					&	"<hr>" & vbCRLF
			END IF
	
		END IF
	
		oRS.Close
		SET oRS = Nothing
	END IF



END FUNCTION

FUNCTION notify_getEvents( oOut, dToday )

	notify_getEvents = ""

	'SET g_oRemindSMTP = oSMTP

	DIM	nDateBegin
	DIM	nDateEnd
	DIM	oCalendar

	nDateBegin = jdateFromVBDate( dToday )
	nDateEnd = nDateBegin + g_UE_nDurationCalendar
	
	'gHtmlOption_encodeEmailAddresses = FALSE
	'g_htmlFormat_pictureFunc = "remindEmailPicture"

	SET oCalendar = loadRemindFiles( nDateBegin, nDateEnd, "email,email2", "holiday,usa,none", FALSE, NOW )
	IF NOT Nothing IS oCalendar THEN
	
		DIM	oXML
		SET oXML = oCalendar.xmldom
		DIM	sTomorrow
		DIM	oEvents
		sTomorrow = "//event[category=""email"" or category=""email2""]"
	
		'IF TRUE THEN
		SET oEvents = oXML.selectSingleNode(sTomorrow)
		IF NOT oEvents IS Nothing THEN
		
			'Load the XSL
			DIM	oXSL
			SET oXSL = remindLoadXmlFile( Server.MapPath("scripts/remind.xslt") )
			
			notify_getEvents = oXML.transformNode(oXSL)
			SET oXSL = Nothing
			SET oEvents = Nothing
		END IF
		SET oXML = Nothing
	END IF
	SET oCalendar = Nothing
	
END FUNCTION




DIM g_cssFile
g_cssFile = ""


FUNCTION notify_getRemindCSS()
	notify_getRemindCSS = ""
	
	DIM	sFile
	sFile = findRemindFile( "remind.css" )
	IF "" <> sFile THEN
		g_cssFile = sFile
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT oFile IS Nothing THEN
			DIM	sText
			sText = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			DIM	i
			i = INSTR( sText, vbCRLF )
			IF 0 < i THEN
				notify_getRemindCSS = MID( sText, i+1 )
			ELSE
				notify_getRemindCSS = sText
			END IF
		END IF
	END IF
END FUNCTION



FUNCTION notify_getEmailNotifyCSS()
	notify_getEmailNotifyCSS = ""

	DIM	sFile
	sFile = Server.MapPath("./scripts/user/emailnotify.css")
	IF g_oFSO.FileExists( sFile ) THEN
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT oFile IS Nothing THEN
			DIM	sText
			sText = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			DIM	i
			i = INSTR( sText, vbCRLF )
			IF 0 < i THEN
				notify_getEmailNotifyCSS = MID( sText, i+1 )
			ELSE
				notify_getEmailNotifyCSS = sText
			END IF
		END IF
	END IF

END FUNCTION

FUNCTION notify_buildStyle( sTextCSS )
	notify_buildStyle = "" _
				&	"<style type=""text/css"">" & vbCRLF _
				&	vbCRLF _
				&	sTextCSS & vbCRLF _
				&	vbCRLF _
				&	vbCRLF _
				&	"</style>" & vbCRLF _
				&	""
END FUNCTION


FUNCTION notify_buildHtmlBody( oOut, dToday )




	' get remind events
	'
	DIM	sEvents
	sEvents = notify_getEvents( oOut, dToday )
	
	
	DIM	sHealthWatch
	'sHealthWatch = processHealthWatch( oOut, dToday )
	sHealthWatch = forum_formatHtml( "forum.xml", g_sDomain, "localforum", g_sSiteTitle & " Forum", 20 )
	
	' get new announcements
	'
	DIM	sAnnouncements
	sAnnouncements = processNotifyAnnouncements( oOut, dToday )
	
	IF "" <> sEvents  OR  "" <> sAnnouncements THEN

		' get CSS
		'
		DIM	sTextRemindCSS
		sTextRemindCSS = notify_getRemindCSS()
		
		DIM	sTextNotifyCSS
		sTextNotifyCSS = notify_getEmailNotifyCSS()
		
		DIM sStyle
		sStyle = notify_buildStyle( sTextRemindCSS & vbCRLF & sTextNotifyCSS )


		DIM	i
		DIM	sBaseURL
		sBaseURL = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")
		i = INSTRREV( sBaseURL, "/" )
		IF 0 < i THEN sBaseURL = LEFT(sBaseURL,i)
		

		DIM	sBody
		
		sBody = "" _
			&	"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCRLF _
			&	"<html>" & vbCRLF _
			&	"<head>" & vbCRLF _
			&	"<base href=""" & sBaseURL & """>" & vbCRLF _
			&	sStyle & vbCRLF _
			&	"</head>" & vbCRLF _
			&	"<body>" & vbCRLF _
			&	"<h2 class=""emailnotification"">" & g_sSiteName & " <i>Upcoming Events</i></h2>" & vbCRLF _
			&	sEvents & vbCRLF _
			&	sHealthWatch & vbCRLF _
			&	sAnnouncements & vbCRLF _
			&	"<p><br></p>" & vbCRLF _
			&	"<p>For more events visit " _
			&		"<a href=""http://www." & g_sDomain & "/"">" _
			&		"www." & g_sDomain _
			&		"</a>" _
			&	"</p>" & vbCRLF _
			&	"<p class=""noreply"">Do not reply to this email</p>" & vbCRLF _
			&	"</body>" & vbCRLF _
			&	"</html>" & vbCRLF


		notify_buildHtmlBody = sBody
		
	END IF

END FUNCTION








%>