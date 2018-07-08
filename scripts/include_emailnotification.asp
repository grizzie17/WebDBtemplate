<!--#include file="mailserver.asp"-->
<!--#include file="mailsend.asp"-->
<%

'
'	depends on include_emailnotificationformat.asp
'

DIM g_subscriptions_sDataFolder

FUNCTION findSubscriptionsFolder()
	DIM	sFolder
	
	IF "" = g_subscriptions_sDataFolder THEN
		sFolder = findAppDataFolder( "emailsubscriptions" )
		IF "" <> sFolder THEN
			findSubscriptionsFolder = sFolder
			g_subscriptions_sDataFolder = sFolder
		ELSE
			findSubscriptionsFolder = ""
		END IF
	ELSE
		findSubscriptionsFolder = g_subscriptions_sDataFolder
	END IF
END FUNCTION


DIM	g_notification_sFile
g_notification_sFile = ""

FUNCTION findNotificationFile()

	IF "" <> g_notification_sFile THEN
		findNotificationFile = g_notification_sFile
	ELSE

		findNotificationFile = ""
		DIM	sFolder
		DIM	sFile
		
		sFile = ""
		sFolder = findSubscriptionsFolder()
		IF "" <> sFolder THEN
			sFile = g_oFSO.BuildPath(sFolder,"lastemailed.txt")
			g_notification_sFile = sFile
			IF g_oFSO.FileExists(sFile) THEN
				findNotificationFile = sFile
			END IF
		END IF
	END IF
	
END FUNCTION


FUNCTION findNotificationErrorFile()

		findNotificationErrorFile = ""
		DIM	sFolder
		DIM	sFile
		
		sFile = ""
		sFolder = findSubscriptionsFolder()
		IF "" <> sFolder THEN
			sFile = g_oFSO.BuildPath(sFolder,"emailerror.txt")
			IF g_oFSO.FileExists(sFile) THEN
				findNotificationErrorFile = sFile
			END IF
		END IF
	
END FUNCTION


SUB writeNotificationErrorFile( sError )
	DIM	sFolder
	DIM	sFile
	
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath( sFolder, "emailerror.txt" )
		IF g_oFSO.FileExists(sFile) THEN
			g_oFSO.DeleteFile(sFile)
		END IF
		DIM	oFile
		SET oFile = g_oFSO.CreateTextFile( sFile, TRUE )
		oFile.WriteLine sError
		oFile.Close
		SET oFile = Nothing
	END IF
END SUB


SUB clearNotificationErrorFile()
	DIM	sFolder
	DIM	sFile
	
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath( sFolder, "emailerror.txt" )
		IF g_oFSO.FileExists(sFile) THEN
			g_oFSO.DeleteFile(sFile)
		END IF
	END IF
END SUB



DIM	g_dSchedNotification

FUNCTION needsNotification()
	needsNotification = FALSE
	
	g_dSchedNotification = CDATE( "1/1/2000" )
	
	DIM	sFile
	sFile = findNotificationFile()
	IF "" = sFile THEN
		needsNotification = TRUE
	ELSEIF "" <> findNotificationErrorFile() THEN
		needsNotification = TRUE
	ELSE
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT Nothing IS oFile THEN
			DIM	sLine
			sLine = oFile.ReadLine
			oFile.Close
			SET oFile = Nothing
			
			DIM	d
			d = CDATE(sLine)
			'Response.Write "<p>scheduled: " & sLine & "<br>"
			'Response.Write "Adjusted Date: " & DATEADD("h", g_nServerTimeZoneOffset, NOW) & "</p>"
			IF 0 <= DATEDIFF( "d", d, DATEADD("h", g_nServerTimeZoneOffset, NOW)) THEN
				g_dSchedNotification = d
				needsNotification = TRUE
				'needsNotification = FALSE
			END IF
		END IF
	END IF
END FUNCTION



SUB	writeNotificationDate()

	DIM	sFile
	sFile = findNotificationFile()
	IF "" <> sFile THEN
		IF g_oFSO.FileExists(sFile) THEN
			g_oFSO.DeleteFile(sFile)
		END IF
		DIM	oFile
		SET oFile = g_oFSO.CreateTextFile( sFile, TRUE )
		
		DIM	d
		d = findNextNotificationDate()
		
		oFile.WriteLine MONTH(d) & "/" & DAY(d) & "/" & YEAR(d)
		oFile.Close
		SET oFile = Nothing
	END IF
END SUB



FUNCTION findSubscriptionsFile()

	findSubscriptionsFile = ""
	DIM	sFolder
	DIM	sFile

	sFile = ""
	sFolder = findSubscriptionsFolder()
	IF "" <> sFolder THEN
		sFile = g_oFSO.BuildPath(sFolder,"subscriptions.txt")
		IF g_oFSO.FileExists(sFile) THEN
			findSubscriptionsFile = sFile
		END IF
	END IF

END FUNCTION






DIM	g_oRemindSMTP
SET g_oRemindSMTP = Nothing


FUNCTION remindEmailPicture( sLabel )

	remindEmailPicture = ""
	IF NOT Nothing IS g_oRemindSMTP THEN

		DIM	sFolder
		DIM	sFile
		sFolder = findRemindFolder()
		sFile = sFolder & "\images\" & sLabel
		IF g_oFSO.FileExists( sFile ) THEN
			DIM	sPictID
			sPictID = g_oRemindSMTP.AddInlineAttachment( sFile )
			IF "" <> sPictID THEN
				remindEmailPicture = "cid:" & sPictID
			END IF
		END IF
	END IF

END FUNCTION



DIM g_oAnnounceSMTP
SET g_oAnnounceSMTP = Nothing

FUNCTION smtp_processNotifyAnnouncementsPicture( sPicture )
	smtp_processNotifyAnnouncementsPicture = ""
	IF NOT g_oAnnounceSMTP IS Nothing THEN
		DIM	sFile
		sFile = picture_buildPath( "pages", sPicture )
		IF g_oFSO.FileExists( sFile ) THEN
			DIM	sPictID
			sPictID = g_oAnnounceSMTP.AddInlineAttachment( sFile )
			IF "" <> sPictID THEN
				smtp_processNotifyAnnouncementsPicture = "cid:" & sPictID
			END IF
		END IF
	END IF
END FUNCTION




FUNCTION notify_setupSMTP()

	SET notify_setupSMTP = Nothing

	DIM oSMTP

	SET oSMTP = mailMakeSMTP()
	IF NOT Nothing IS oSMTP THEN
		oSMTP.Server = mailServer()
		
		DIM	sAccount
		DIM	sPW
		
		sAccount = g_sAnnonUser
		sPW = g_sAnnonPW
		
		IF "" <> sAccount  AND  "" <> sPW THEN
			oSMTP.User = sAccount
			oSMTP.PW = sPW
		END IF
		
		
		
		oSMTP.From = "Notifications@" & g_sDomain
		IF "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
			oSMTP.FromName = "Event Notification"
			oSMTP.AddRecipient2 "UpcomingEvents@" & g_sDomain, "Upcoming Events"
			'oSMTP.AddRecipientBCC "UpcomingEvents@" & g_sDomain
		ELSE
			oSMTP.FromName = "Local " & g_sDomain
			oSMTP.AddRecipientBCC "Webmaster@" & g_sDomain
		END IF
		SET notify_setupSMTP = oSMTP
	END IF
	SET oSMTP = Nothing

END FUNCTION







FUNCTION processNotification()

	processNotification = ""

	DIM	oSMTP
	
	SET oSMTP = notify_setupSMTP()
	IF NOT Nothing IS oSMTP THEN
	
		DIM	sSavePictureFunc
		DIM	sSaveAnnounceFunc

		SET g_oRemindSMTP = oSMTP
		sSavePictureFunc = g_htmlFormat_pictureFunc
		g_htmlFormat_pictureFunc = "remindEmailPicture"
		g_emailnotify_processHtmlPicture = "notifySMTP_f_processHtmlPicture"
		gHtmlOption_encodeEmailAddresses = FALSE
		SET g_oAnnounceSMTP = oSMTP
		sSaveAnnounceFunc = g_sProcessAnnouncementsPictureFunc
		g_sProcessAnnouncementsPictureFunc = "smtp_processNotifyAnnouncementsPicture"

		DIM	dToday
		dToday = DATEADD("h", g_nServerTimeZoneOffset, NOW)
		
		DIM sBody
		sBody = notify_buildHtmlBody( oSMTP, dToday )
		
		IF "" <> sBody THEN			
	
			oSMTP.Subject = g_sSiteName & " upcoming events"
			
			oSMTP.Body = sBody
					

			'send it
			DIM	sError
			sError = oSMTP.Send
			IF "0:" = LEFT(sError,2) THEN
				clearNotificationErrorFile
				writeNotificationDate
				processNotification = FORMATDATETIME(dToday, 0)
			ELSE
				writeNotificationErrorFile sError
			END IF

		ELSE
			clearNotificationErrorFile
			writeNotificationDate
		END IF

		SET g_oRemindSMTP = Nothing
		g_htmlFormat_pictureFunc = sSavePictureFunc
		SET g_oAnnounceSMTP = Nothing
		g_sProcessAnnouncementsPictureFunc = sSaveAnnounceFunc
		
	END IF
	SET oSMTP = Nothing

END FUNCTION







SUB emailNotification()

	IF "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
		IF needsNotification() THEN
			'Response.Write "<p>Needs Notification</p>" & vbCRLF
			'writeNotificationDate
			processNotification
		END IF
	END IF

END SUB




'emailNotification


%>
