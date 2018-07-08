<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit

Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%





	DIM	sField
	DIM	sValue
	DIM	nType
	DIM	sDataValue
	DIM	sSelect
	DIM	nMaximum
	DIM	nSize
	DIM	nRID
	DIM	oRS
	DIM	oRSr
	DIM	n

	DIM	aFields(2,50)
	n = 0
	FOR EACH sField IN Request.Form
		IF "B1" <> sField THEN
			sValue = TRIM(Request.Form(sField))

			sSelect = "" _
			&	"SELECT * " _
			&	"FROM " & dbQ("config") & " " _
			&	"WHERE " & dbQ("Name") & "='" & sField & "' " _
			&	";"
			SET oRSr = dbQueryRead( g_DC, sSelect )
			IF NOT oRSr IS Nothing THEN
				oRSr.MoveFirst
				nRID = recNumber(oRSr.Fields("RID"))
				nMaximum = recNumber(oRSr.Fields("Maximum"))
				nType = recNumber(oRSr.Fields("Type"))
				sDataValue = recString(oRSr.Fields("DataValue"))
				nSize = oRSr.Fields("DataValue").DefinedSize
				'Response.Write " --- RID: " & nRID & ", Max: " & nMaximum & ", Type: " & nType & "<br>" & vbCRLF
				oRSr.ActiveConnection = Nothing
				oRSr.Close
				SET oRSr = Nothing
				SELECT CASE nType
				CASE 0	'string
					IF 0 < nMaximum THEN
						IF nMaximum < LEN(sValue) THEN
							sValue = LEFT(sValue,nMaximum)
						END IF
					END IF
				CASE 1	'integer
					IF ISNUMERIC(sValue) THEN
						IF nMaximum < CLNG(sValue) THEN
							sValue = CSTR(nMaximum)
						END IF
					END IF
				CASE 2	'bool
					nSize = 5
					IF 0 < LEN(sValue) THEN
						sValue = "true"
					ELSE
						sValue = "false"
					END IF
				CASE ELSE
				END SELECT
				IF nSize < LEN(sValue) THEN
					sValue = LEFT(sValue,nSize)
				END IF
				IF sValue <> sDataValue THEN
					aFields(0,n) = sField
					aFields(1,n) = fieldString(sValue)
					Response.Write " " & aFields(0,n) & " = " & aFields(1,n) & "<br>" & vbCRLF
					n = n + 1
				END IF
			END IF
		END IF
	NEXT 'sField

	DIM i
	FOR i = 0 TO n-1
		sSelect = "" _
		&	"SELECT * " _
		&	"FROM config " _
		&	"WHERE Name = '" & aFields(0,i) & "' " _
		&	";"
		SET oRS = dbQueryUpdate( g_DC, sSelect )
		IF NOT oRS IS Nothing THEN
			oRS.MoveFirst
			IF NOT oRS.EOF THEN
				oRS.Fields("DataValue").Value = aFields(1,i)
				oRS.Update
			END IF
			oRS.Close
			SET oRS = Nothing
		END IF
	NEXT 'i



cache_clearFolder "site"

%>
<!--#include file="scripts\include_db_end.asp"-->
<%

Response.Redirect "./"

%>