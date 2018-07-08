<%@ Language=VBScript %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")


DIM	g_sPage
g_sPage = Request("cat")

DIM	sMon
DIM	m
sMon = Request("m")
IF 0 < LEN(sMon) THEN
	m = CINT(sMon)
ELSE
	m = 3
END IF

DIM g_sNavigateTabLabel
g_sNavigateTabLabel = "(tab|special)"


%>
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts/remind.asp"-->
<!--#include file="scripts/remind_files.asp"-->
<!--#include file="scripts/remind_utils.asp"-->
<%

Response.ContentType="text/plain"

'===========================================================



'===========================================================



FUNCTION remind_handleSubsTextDummy( sText, nDate, nTime )
	remind_handleSubsTextDummy = htmlFormatEncodeText( sText )
END FUNCTION

FUNCTION dummy_buildLink( sProtocol, sURL, sTarget, sText )
	DIM	sLink
	sLink = sProtocol & sURL
	IF sLink = sText THEN
		IF "address:" = sProtocol THEN
			dummy_buildLink = REPLACE(sURL, "+", " ")
		ELSEIF "mailto:" = sProtocol THEN
			dummy_buildLink = sURL
		ELSE
			dummy_buildLink = sLink
		END IF
	ELSE
		dummy_buildLink = sText
	END IF
END FUNCTION

FUNCTION dummyPictureGen( sAlign, sBorder, sWidth, sColor, sCaption, sFile )
	dummyPictureGen = " "
END FUNCTION


FUNCTION dummyPictureFile( sLabel )
	dummyPictureFile = " "
END FUNCTION



FUNCTION getStyleParam( sName, sParams )
	getStyleParam = ""
	DIM	oReg
	SET oReg = NEW RegExp
	oReg.Pattern = "\s*" & sName & ":\s*([^;]+);"
	oReg.Global = TRUE
	oReg.IgnoreCase = TRUE
	DIM	oMatchList
	DIM	oMatch
	SET oMatchList = oReg.Execute( sParams )
	IF NOT oMatchList IS Nothing THEN
		FOR EACH oMatch IN oMatchList
			getStyleParam = TRIM(oMatch.Submatches(0))
			EXIT FOR
		NEXT
		SET oMatchList = Nothing
	END IF
	SET oReg = Nothing
END FUNCTION



CLASS OLEvent

	DIM	m_sSubject
	DIM	m_sDescription
	DIM	m_sLocation
	DIM	m_sCategories
	DIM	m_dStartDate
	DIM	m_tStartTime
	DIM	m_dEndDate
	DIM	m_tEndTime
	DIM	m_bAllDayEvent
	DIM	m_bReminderOnOff
	DIM	m_dReminderDate
	DIM	m_tReminderTime
	DIM	m_sMeetingOrganizer
	DIM	m_sRequiredAttendees
	DIM	m_sOptionalAttendees
	DIM	m_sMeetingResources
	DIM	m_sBillingInformation
	DIM	m_fMileage
	DIM	m_sPriority
	DIM	m_bPrivate
	DIM	m_sSensitivity
	DIM	m_nShowTimeAs

	SUB Class_Initialize()
	END SUB

	SUB Class_Terminate()
	END SUB
	

END CLASS







g_htmlFormat_makeLinkFunc = "dummy_buildLink"

'g_remind_handleSubsText = "remind_handleSubsTextDummy"



DIM	oCalendar


DIM	nYear, nYearEnd
DIM	nMon, nMonEnd

nYear = YEAR(NOW)
nMon = MONTH(NOW)
nMonEnd = nMon + m
IF 12 < nMonEnd THEN
	nYearEnd = nYear + FIX(nMonEnd/12)
	nMonEnd = nMonEnd MOD 12
ELSE
	nYearEnd = nYear
END IF

'Response.Write "m = " & nMon & ", y = " & nYear & "<br>"
'Response.Write "m = " & nMonEnd & ", y = " & nYearEnd & "<br>" & vbCRLF
		
DIM	nDateBegin
DIM	nDateEnd

nDateBegin = jdateFromGregorian( 1, nMon, nYear )
nDateEnd = jdateFromGregorian( 1, nMonEnd, nYearEnd ) - 1



'nDateBegin = jdateFromVBDate( NOW )
'nDateEnd = nDateBegin + 28 * 3 - 1
'nDateEnd = -14	'pending

gHtmlOption_encodeEmailAddresses = FALSE

DIM	sSavePictureFunc
sSavePictureFunc = g_htmlFormat_pictureFunc
g_htmlFormat_pictureFunc = "dummyPictureFile"
g_htmlFormat_pictureGenFunc = "dummyPictureGen"


SET oCalendar = loadRemindFiles2( nDateBegin, nDateEnd, _
						g_sPage, "holiday,usa,none", FALSE, NOW, "list" )
IF NOT Nothing IS oCalendar THEN

	'Load the XSL
	DIM	oXSL
	DIM	oXML
	SET oXSL = remindLoadXmlFile( Server.MapPath("scripts/remind_outlook.xslt") )

	SET oXML = oCalendar.xmldom
	DIM	sText
	sText = oXML.transformNode(oXSL)
	sText = REPLACE(sText, "&#8211;", "-" )
	sText = REPLACE(sText, "&#8217;", "'" )
	sText = REPLACE(sText, "&#8216;", "'" )
	sText = REPLACE(sText, "&#8220;", """" )
	sText = REPLACE(sText, "&#8221;", """" )
	sText = REPLACE(sText, "&nbsp;", " " )
	sText = REPLACE(sText, "&amp;", "&" )
	sText = REPLACE(sText, "&quot;", """" )
	sText = REPLACE(sText, "&lt;", "<" )
	sText = REPLACE(sText, "&gt;", ">" )
	'sText = REPLACE(sText, "", "" )



	DIM	sCSSFile
	DIM	sCSSText
	sCSSFile = findRemindFile( "remind.css" )
	IF "" <> sCSSFile THEN
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile( sCSSFile, 1 )
		IF NOT oFile IS Nothing THEN
			sCSSText = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing


			DIM	oReg
			DIM	oMatchList
			DIM	oMatch
			DIM	sClass
			DIM	sParams
			DIM	xP
			DIM	s
			xP = ""


			SET oReg = NEW RegExp
			oReg.Pattern = "\s\.([\w_]+)\s*\{\s*([^\}]+)\}"
			oReg.IgnoreCase = TRUE
			oReg.Global = TRUE
			SET oMatchList = oReg.Execute( sCSSText )
			IF NOT oMatchList IS Nothing THEN
				FOR EACH oMatch IN oMatchList
					sClass = oMatch.Submatches(0)
					sParams = oMatch.Submatches(1)
					IF "" <> sParams THEN
						xP = ""
						s = getStyleParam( "color", sParams )
						xP = s
						s = getStyleParam( "font-weight", sParams )
						IF "bold" = LCASE(s) THEN xP = xP & " b"
						s = getStyleParam( "font-size", sParams )
						SELECT CASE LCASE(s)
						CASE "medium"
							xP = xP & " medium"
						CASE "small"
							xP = xP & " small"
						CASE "x-small"
							xP = xP & " x-small"
						CASE "xx-small"
							xP = xP & " xx-small"
						CASE "large"
							xP = xP & " large"
						CASE "x-large"
							xP = xP & " x-large"
						CASE "xx-large"
							xP = xP & " xx-large"
						CASE "larger"
							xP = xP & " larger"
						CASE "smaller"
							xP = xP & " smaller"
						END SELECT
						s = getStyleParam( "font-style", sParams )
						IF "italic" = LCASE(s) THEN xP = xP & " i"
						s = getStyleParam( "text-decoration", sParams )
						IF 0 < INSTR(LCASE(s), "underline") THEN xP = xP & " u"
						IF 0 < INSTR(LCASE(s), "line-through") THEN xP = xP & " strike"
						s = getStyleParam( "font-family", sParams )
						IF "" <> s THEN
							s = htmlFormatMetaGenericFontName( s )
							IF "" <> s THEN xP = xP & " " & s
						END IF
						xP = TRIM(xP)
						IF 0 < INSTR(xP, " " ) THEN
							xP = "style " & xP
						END IF
						sText = REPLACE(sText,"~~"&sClass&"~~", xP)
					END IF
				NEXT 'oMatch
				SET oMatchList = Nothing
			END IF
			SET oReg = Nothing


		END IF
	END IF



	Response.Write sText
	
	SET oXML = Nothing
	SET oXSL = Nothing

END IF

SET g_oFSO = Nothing


%> 
