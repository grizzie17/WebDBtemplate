<!--#include file="ado.asp"-->
<%



FUNCTION access_dbSessionKey( sDBName )
	DIM	i
	DIM	sSessionKey
	sSessionKey = "MDBFile" & sDBName & Request.ServerVariables("PATH_INFO")
	i = INSTRREV(sSessionKey,"/")
	IF 0 < i THEN sSessionKey = LEFT(sSessionKey,i-1)
	access_dbSessionKey = sSessionKey
END FUNCTION


FUNCTION access_dbConnect( sDBName )

	SET access_dbConnect = Nothing
	
	DIM	sSessionKey
	sSessionKey = access_dbSessionKey( sDBName )
	
	DIM	sDBFile
	sDBFile = Session.Contents(sSessionKey)
	IF "" = sDBFile THEN
		sDBFile = findDBFile( sDBName & ".mdb" )
	END IF
	
	DIM	oDC
	SET oDC = Nothing
	IF "" <> sDBFile THEN
		Session(sSessionKey) = sDBFile
		SET oDC = Server.CreateObject( "ADODB.Connection" )
		IF NOT Nothing IS oDC THEN
			DIM	sConn
			'sConn = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & sDBFile & ";"
			sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBFile & ";"
			ON ERROR RESUME NEXT
			oDC.open sConn
			IF Err THEN
				SET oDC = Nothing
			END IF
			ON ERROR GOTO 0
			SET access_dbConnect = oDC
			SET oDC = Nothing
		END IF
	END IF

END FUNCTION



FUNCTION access_dbQueryRead( oDC, sQuery )

	SET access_dbQueryRead = Nothing
	IF NOT Nothing IS oDC THEN
		DIM	oRS
		SET oRS = Server.CreateObject("ADODB.RecordSet")
		IF NOT Nothing IS oRS THEN
			DIM	nErr
			With oRS
				.CursorLocation = adUseClient
				.CursorType     = adOpenStatic 'adOpenKeyset
				.LockType       = adLockReadOnly
				ON ERROR Resume Next
					.Open sQuery, oDC
					nErr = Err
				ON ERROR GOTO 0
				IF 0 <> nErr THEN
					Response.Write "access_dbQueryRead -- error in query<br>" & vbCRLF
					Response.Write "<p>" & Server.HTMLEncode(sQuery) & "</p>" & vbCRLF
					Response.Flush
					SET oRS = Nothing
				ELSE
					Set .ActiveConnection = Nothing
					IF oRS.EOF THEN
						oRS.Close
						SET oRS = Nothing
					END IF
				END IF
			End With
		ELSE
			Response.Write "access_dbQueryRead -- error creating ADODB.RecordSet<br>" & vbCRLF
		END IF
		SET access_dbQueryRead = oRS
		SET oRS = Nothing
	END IF

END FUNCTION


FUNCTION access_dbQueryUpdate( oDC, sQuery )

	SET access_dbQueryUpdate = Nothing
	

END FUNCTION











FUNCTION access_recBool( o )
	access_recBool = FALSE
	IF NOT ISNULL( o ) THEN
		IF CBOOL( o.Value ) THEN
			access_recBool = TRUE
		ELSE
			access_recBool = FALSE
		END IF
	END IF
END FUNCTION


FUNCTION access_isTrue( o )
	access_isTrue = access_recBool( o )
END FUNCTION





FUNCTION access_recString( o )
	IF ISNULL( o ) THEN
		access_recString = ""
	ELSEIF ISNULL( o.Value ) THEN
		access_recString = ""
	ELSE
		access_recString = TRIM(CSTR(o.Value))
	END IF
END FUNCTION


FUNCTION access_recNumber( o )
	IF ISNULL( o ) THEN
		access_recNumber = 0
	ELSEIF ISNULL( o.Value ) THEN
		access_recNumber = 0
	ELSEIF ISNUMERIC( o.Value ) THEN
		access_recNumber = CLNG(o.Value)
	ELSE
		access_recNumber = 0
	END IF
END FUNCTION


FUNCTION access_recCurrency( o )
	IF ISNULL( o ) THEN
		access_recCurrency = 0.00
	ELSEIF ISNULL( o.Value ) THEN
		access_recCurrency = 0.00
	ELSEIF ISNUMERIC( o.Value ) THEN
		access_recCurrency = CCUR(o.Value)
	END IF
END FUNCTION


FUNCTION access_recDate( o )
	IF ISNULL( o ) THEN
		access_recDate = CDATE( "12:00:00AM" )
	ELSEIF ISNULL( o.Value ) THEN
		access_recDate = CDATE( "12:00:00AM" )
	ELSEIF ISDATE( o.Value ) THEN
		access_recDate = CDATE( o.Value )
	END IF
END FUNCTION







%>