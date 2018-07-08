<%



FUNCTION recConfigString( oRS, sField, sDefault )
	DIM	oField
	SET oField = oRS.Fields(sField)
	DIM	sTemp
	sTemp = recString( oField )
	IF "" <> sTemp THEN
		recConfigString = sTemp
	ELSE
		recConfigString = sDefault
	END IF
	SET oField = Nothing
	
END FUNCTION

FUNCTION recConfigNumber( oRS, sField, nDefault )
	DIM	oField
	SET oField = oRS.Fields(sField)
	IF ISNULL(oField) THEN
		recConfigNumber = nDefault
	ELSE
		recConfigNumber = recNumber( oField )
	END IF
	SET oField = Nothing
END FUNCTION

FUNCTION recConfigBool( oRS, sField, nDefault )
	DIM	oField
	SET oField = oRS.Fields(sField)
	IF ISNULL(oField) THEN
		recConfigBool = nDefault
	ELSE
		recConfigBool = recBool( oField )
	END IF
	SET oField = Nothing
END FUNCTION


FUNCTION configdb_buildString()

	DIM	sConfig
	sConfig = ""

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		" * " _
		&	"FROM " _
		&		dbQ("config") & " " _
		&	";"
	
	DIM	oRS
	SET oRS = dbQueryRead( g_DC, sSelect )
	
	IF NOT oRS IS Nothing THEN
	
		IF NOT oRS.EOF THEN

			DIM	xQ
			xQ = """"
			DIM	xNL
			xNL = vbCRLF
			DIM	xQNL
			xQNL = xQ & xNL

			DIM	o
			DIM	sTemp
			DIM	nTemp
			DIM	sName
			DIM	sPrefix
			DIM	sValue
			DIM	nValue
			DIM	bValue

			oRS.MoveFirst
			DO UNTIL oRS.EOF
				sName = recString( oRS.Fields("Name") )
				IF 0 < LEN(sName) THEN
					SELECT CASE oRS.Fields("Type").Value
					CASE 0	'string
						sPrefix = "g_s"
						sValue = recConfigString( oRS, "DataValue", EVAL(sPrefix & sName) )
						sConfig = sConfig & sPrefix & sName & " = " & xQ & sValue & xQNL
					CASE 1	'integer
						sPrefix = "g_n"
						nValue = recConfigNumber( oRS, "DataValue", EVAL(sPrefix & sName) )
						sConfig = sConfig & sPrefix & sName & " = " & nValue & xNL
					CASE 2	'bool
						sPrefix = "g_b"
						bValue = recConfigBool( oRS, "DataValue", EVAL(sPrefix & sName) )
						sConfig = sConfig & sPrefix & sName & " = " & bValue & xNL
					CASE ELSE
					END SELECT
				END IF
				oRS.MoveNext
			LOOP
		
		ELSE
			Response.Write "configdb -- oRS returned EOF<br>" & vbCRLF

		END IF
	
		oRS.Close
		SET oRS = Nothing
	ELSE
		Response.Write "configdb -- dbQueryRead failed<br>" & vbCRLF
	
	END IF

	configdb_buildString = sConfig

END FUNCTION


SUB getConfigData()

	DIM	sConfig
	DIM	oFile
	SET oFile = cache_openTextFile( "site", "config.asp", "d", 5, "w" )
	IF NOT oFile IS Nothing THEN
		sConfig = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sConfig = configdb_buildString()
		SET oFile = cache_makeFile( "site", "config.asp" )
		IF NOT oFile IS Nothing THEN
			oFile.Write sConfig
			oFile.Close
			SET oFile = Nothing
		END IF
	END IF
	
	EXECUTEGLOBAL sConfig

	IF 0 < g_nCopyrightStartYear THEN
		g_sCopyright = "&copy; " & g_nCopyrightStartYear & " .. " & DatePart("yyyy",NOW) & "&nbsp;  " & g_sShortSiteName & "&nbsp; All rights reserved"
	ELSE
		g_sCopyright = "&copy; " & DatePart("yyyy",NOW) & "&nbsp;  " & g_sShortSiteName & "&nbsp; All rights reserved"
	END IF


END SUB
getConfigData


IF "" = g_sCalendarList THEN

g_sCalendarList = "" _
	&	"All" & vbTAB & "" & vbTAB & "Events" & vbLF _
	&	UCASE(g_sSiteChapter) & " Events" & vbTAB & LCASE(g_sSiteChapter) & vbTAB & UCASE(g_sSiteChapter) & " Events" & vbLF _
	&	"Activities" & vbTAB & "activity" & vbTAB & "Activities" & vbLF _
	&	"Charity" & vbTAB & "charity" & vbTAB & "Charity Events / Rides" & vbLF _
	&	"Education" & vbTAB & "safety" & vbTAB & "Education" & vbLF _
	&	"Meetings" & vbTAB & "meeting" & vbTAB & "Meetings" & vbLF _
	&	"Rallies" & vbTAB & "rally" & vbTAB & "Rallies" & vbLF _
	&	"Trips" & vbTAB & "trip" & vbTAB & "Trips" & vbLF _
	&	"Birthdays" & vbTAB & "birthday,anniversary" & vbTAB & "Birthdays"

END IF

IF "" = g_sCalendarHiddenList THEN

g_sCalendarHiddenList = "" _
	&	"Email" & vbTAB & "email,email2" & vbTAB & "Email Events" & vbLF _
	&	"Newsletter" & vbTAB & "key,newsletter," & LCASE(g_sSiteChapter) & ",rally,district,funday,safety" & vbTAB & "Newsletter Events" & vbLF _
	&	"Ride Meeting" & vbTAB & "key,email," & LCASE(g_sSiteChapter) & ",rally,district,funday,safety,entertainment,charity" & vbTAB & "Ride Planning"

END IF

%>