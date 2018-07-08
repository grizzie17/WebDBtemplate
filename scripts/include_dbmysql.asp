<!--#include file="ado.asp"-->
<!--#include file="findfiles.asp"-->
<%

FUNCTION dbVendor()
	dbVendor = "mysql"
END FUNCTION


FUNCTION dbSessionKey( sDBName )
	DIM	i
	DIM	sSessionKey
	sSessionKey = "MDBFile" & sDBName & Request.ServerVariables("PATH_INFO")
	i = INSTRREV(sSessionKey,"/")
	IF 0 < i THEN sSessionKey = LEFT(sSessionKey,i-1)
	dbSessionKey = sSessionKey
END FUNCTION


SUB loadUserConfig()
	DIM	sPath
	sPath = findFileUpTree( "scripts\user\configuser.asp" )
	IF 0 < LEN(sPath) THEN
		DIM	oFile
		SET oFile = g_oFSO.OpenTextFile(sPath)
		IF NOT oFile IS Nothing THEN
			DIM	sConfig
			sConfig = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			DIM	i
			' get rid of the asp brackets
			i = INSTR(sConfig, "<"&"%")
			IF -1 < i THEN
				sConfig = MID(sConfig, i+2 )
				sConfig = REPLACE( sConfig, "%"&">", "", 1, -1 )
			END IF
			EXECUTEGLOBAL sConfig
		END IF
	END IF
END SUB


FUNCTION dbConnectAuth( sDBName, sUser, sPassword )

	SET dbConnectAuth = Nothing

	loadUserConfig
	
	DIM	oDC
	SET oDC = Nothing
	SET oDC = Server.CreateObject( "ADODB.Connection" )
	IF NOT Nothing IS oDC THEN
		DIM	sConn
		sConn = "DSN=grizzie_" & sDBName & ";" _
			&	"Uid=" & sUser & ";" _
			&	"Pwd=" & sPassword & ";" _
			&	"Option=(1 + 2 + 8 + 32 + 16384);"
			'     1 = client can't handle the real column width being returned
			'     2 = have MySQL return found rows value
			'     8 = Allow big values: required for BLOB support
			'    32 = Dynamic cursor support
			' 16384 = Convert LongLong into Int: allows large numeric results to get properly interpreted
		ON ERROR RESUME NEXT
			oDC.open sConn
			DIM	sErr, nErr
			nErr = Err.number
			sErr = Err.Description
		ON ERROR GOTO 0
		IF 0 <> nErr THEN
			Response.Write "dbConnectAuth -- error connecting<br>" & vbCRLF
			Response.Write " -- Error = " & CSTR(HEX(nErr)) & " - " & sErr & "<br>" & vbCRLF
			Response.Write "<p>" & Server.HTMLEncode(sQuery) & "</p>" & vbCRLF
			Response.Flush
		ELSE
			IF oDC.state <> adStateOpen THEN
				DIM	aErrors
				SET aErrors = oDC.Errors
				DIM oErr
				FOR EACH oErr in aErrors
					Response.Write "Error = " & oErr.description & " (" & CSTR(HEX(oErr.number)) & ")<br>" & vbCRLF
				NEXT
				Response.Write "Error opening database = " & CSTR(HEX(Err)) & "<br>" & vbCRLF
			ELSE
				oDC.cursorLocation = adUseClient
				SET dbConnectAuth = oDC
			END IF
		END IF
		SET oDC = Nothing
	ELSE
		Response.Write "Error creating ADODB.Connection<br>" & vbCRLF
	END IF

END FUNCTION

FUNCTION dbConnect( sDBName )
	SET dbConnect = dbConnectAuth( sDBName, g_sDBUser, g_sDBPassword )
END FUNCTION



FUNCTION dbXLTQuery( sQuery )
	DIM	s
	s = sQuery
	s = REPLACE(REPLACE(s,"[","`"),"]","`")
	s = REPLACE(s, "GETDATE(", "CURDATE(" )
	s = REPLACE(s, "ISNULL(", "IFNULL(" )
	s = REPLACE(s, "DATEDIFF(day,", "DATEDIFF(" )
	dbXLTQuery = s
END FUNCTION



FUNCTION dbQueryRead( oDC, sQuery )

	SET dbQueryRead = Nothing
	IF NOT Nothing IS oDC THEN
		DIM	oRS
		SET oRS = Server.CreateObject("ADODB.RecordSet")
		IF NOT Nothing IS oRS THEN
			With oRS
				.CursorLocation = adUseClient
				.CursorType     = adOpenStatic 'adOpenKeyset
				.LockType       = adLockReadOnly
			End With
			ON ERROR Resume Next
				oRS.Open dbXLTQuery(sQuery), oDC
				DIM	nErr, sErr
				nErr = Err.number
				sErr = Err.Description
			ON ERROR GOTO 0
			IF 0 <> nErr THEN
				Response.Write "dbQueryRead -- error in query<br>" & vbCRLF
				Response.Write " -- Error = " & CSTR(HEX(nErr)) & " - " & sErr & "<br>" & vbCRLF
				Response.Write "<p>" & Server.HTMLEncode(sQuery) & "</p>" & vbCRLF
				Response.Flush
				SET oRS = Nothing
			ELSE
				Set oRS.ActiveConnection = Nothing
			END IF
			SET dbQueryRead = oRS
			SET oRS = Nothing
		ELSE
			Response.Write "dbQueryRead -- error creating ADODB.RecordSet<br>" & vbCRLF
		END IF
	ELSE
		Response.Write "dbQueryRead -- missing DataConnection<br>" & vbCRLF
	END IF

END FUNCTION


FUNCTION dbQueryReadEx( oDC, sQuery )
	SET dbQueryReadEx = dbQueryRead( oDC, sQuery )
	IF NOT dbQueryReadEx IS Nothing THEN
		IF dbQueryReadEx.EOF THEN
			dbQueryReadEx.Close
			SET dbQueryReadEx = Nothing
		END IF
	END IF
END FUNCTION


FUNCTION dbQueryUpdate( oDC, sQuery )

	SET dbQueryUpdate = Nothing
	
	IF NOT Nothing IS oDC THEN

		DIM	oRS
		SET oRS = Server.CreateObject("ADODB.RecordSet")
		IF NOT Nothing IS oRS THEN
			With oRS
				.CursorLocation = adUseClient
				.CursorType     = adOpenDynamic
				.LockType       = adLockOptimistic
				.Open dbXLTQuery(sQuery), oDC
			End With
		END IF
		SET dbQueryUpdate = oRS
		SET oRS = Nothing
	END IF

END FUNCTION




FUNCTION dbExecute( oDC, sQuery )

	SET dbExecute = Nothing

	DIM	oRS
	SET oRS = oDC.Execute( dbXLTQuery(sQuery) )
	SET dbExecute = oRS
	SET oRS = Nothing

END FUNCTION


FUNCTION dbGetFieldSize( oDC, sTableName, sFieldName )

	dbGetFieldSize = 0

	DIM	sSelect
	sSelect = "" _
		&	"SELECT " _
		&		"CHARACTER_MAXIMUM_LENGTH " _
		&	"FROM " _
		&		"INFORMATION_SCHEMA.COLUMNS " _
		&	"WHERE " _
		&		"TABLE_NAME = '" &sTableName& "' AND COLUMN_NAME = '" & sFieldName & "' " _
		&	";"
	DIM	oRS
	SET oRS = dbQueryReadEx( g_DC, sSelect )
	IF NOT oRS IS Nothing THEN
		dbGetFieldSize = CINT(oRS.Fields("CHARACTER_MAXIMUM_LENGTH").Value)
		oRS.Close
		SET oRS = Nothing
	END IF

END FUNCTION


FUNCTION dbQ( sName )
	dbQ = "`" & sName & "`"
END FUNCTION


FUNCTION dbT( sName )
	dbT = "'" & sName & "'"
END FUNCTION








FUNCTION recBool( o )
	recBool = FALSE
	IF NOT ISNULL( o ) THEN
		IF CBOOL( o.Value ) THEN
			recBool = TRUE
		ELSE
			recBool = FALSE
		END IF
	END IF
END FUNCTION


FUNCTION isTrue( o )
	isTrue = recBool( o )
END FUNCTION





FUNCTION recString( o )
	IF ISNULL( o ) THEN
		recString = ""
	ELSEIF ISNULL( o.Value ) THEN
		recString = ""
	ELSE
		recString = TRIM(CSTR(o.Value))
	END IF
END FUNCTION


FUNCTION recNumber( o )
	IF ISNULL(o) THEN
		recNumber = 0
	ELSE
		recNumber = CLNG(o.Value)
	END IF
END FUNCTION


FUNCTION recCurrency( o )
	IF ISNULL( o ) THEN
		recCurrency = 0.00
	ELSEIF ISNULL( o.Value ) THEN
		recCurrency = 0.00
	ELSEIF ISNUMERIC( o.Value ) THEN
		recCurrency = CCUR(o.Value)
	END IF
END FUNCTION


FUNCTION recDate( o )
	IF ISNULL( o ) THEN
		recDate = CDATE( "12:00:00AM" )
	ELSEIF ISNULL( o.Value ) THEN
		recDate = CDATE( "12:00:00AM" )
	ELSEIF ISDATE( o.Value ) THEN
		recDate = CDATE( o.Value )
	END IF
END FUNCTION


FUNCTION recBlob( o )
	IF ISNULL( o ) THEN
		SET recBlob = Nothing
	ELSEIF ISNULL( o.Value ) THEN
		SET recBlob = Nothing
	ELSE
		DIM	oStream
		SET oStream = Server.CreateObject("ADODB.Stream")
		IF NOT oStream IS Nothing THEN
			oStream.Write o.Value
			SET recBlob = oStream
			SET oStream = Nothing
		ELSE
			SET recBlob = Nothing
		END IF
	END IF
END FUNCTION





FUNCTION fieldString( s )
	IF "" = s THEN
		fieldString = NULL
	ELSE
		fieldString = CSTR(s)
	END IF
END FUNCTION


FUNCTION fieldBool( s )
	fieldBool = 0
	SELECT CASE LCASE(s)
	CASE "yes", "true", "on"
		fieldBool = -1
	CASE "no", "false", "off"
		fieldBool = 0
	CASE ELSE
		IF "" <> s THEN
			IF ISNUMERIC(s) THEN
				DIM	n
				n = CLNG(s)
				IF 0 <> n THEN
					fieldBool = -1
				END IF
			END IF
		END IF
	END SELECT
END FUNCTION



FUNCTION fieldNumber( s )
	IF "" = CSTR(s) THEN
		fieldNumber = NULL
	ELSEIF ISNUMERIC( s ) THEN
		fieldNumber = s
	ELSE
		fieldNumber = 0
	END IF
END FUNCTION


FUNCTION fieldDate( s )
	IF "" = CSTR(s) THEN
		fieldDate = NULL
	ELSEIF ISDATE( s ) THEN
		IF CDATE(s) = CDATE("12:00:00AM") THEN
			fieldDate = NULL
		ELSE
			fieldDate = s
		END IF
	ELSE
		fieldDate = s
	END IF
END FUNCTION

' input is an ADODB.Stream
FUNCTION fieldBlob( o )
	IF ISNULL( o ) THEN
	ELSE
	END IF
END FUNCTION



%>