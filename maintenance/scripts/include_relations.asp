<%




SUB relations_getArrayFromString( aArray(), sRelations )

	DIM	sList
	DIM	aList
	DIM	aTemp
	DIM	sRel
	DIM	i
	DIM	nJTS
	DIM	nPG
	
	sList = REPLACE( sRelations, vbCRLF, vbLF, 1, -1, vbTextCompare )
	sList = REPLACE( sList, vbCR, vbLF, 1, -1, vbTextCompare )
	sList = REPLACE( sList, ",", vbLF, 1, -1, vbTextCompare )
	aList = SPLIT( sList, vbLF )
	
	REDIM aArray(UBOUND(aList))
	
	FOR i = 0 TO UBOUND(aList)
		sRel = TRIM(aList(i))
		IF "" <> sRel THEN
			aArray(i) = CLNG(sRel)
		ELSE
			aArray(i) = 0
		END IF
	NEXT 'i
	
END SUB


FUNCTION relations_getStringFromArray( aArray() )
	DIM	sList
	DIM	i
	sList = ""
	FOR i = 0 TO UBOUND(aArray)
		IF 0 < aArray(i) THEN
			sList = sList & vbLF & aArray(i)
		END IF
	NEXT 'i
	relations_getStringFromArray = MID(sList,2)
END FUNCTION


SUB	relations_deleteByQuery( sQuery )
	DIM	oRS
	SET oRS = dbQueryUpdate( g_DC, sQuery )
	
	IF NOT oRS.EOF THEN
	
		oRS.MoveFirst		
		DO UNTIL oRS.EOF
			oRS.Delete 1 'adAffectCurrent
			oRS.MoveNext
		LOOP
	END IF
	
	Set oRS.ActiveConnection = Nothing
	oRS.Close
	SET oRS = Nothing
END SUB


SUB relations_deleteTabs( n )

	DIM	sSelect
	sSelect = "SELECT * " _
			&	"FROM tabcategorymap " _
			&	"WHERE " _
			&	" CategoryID = " & n & ";"
			
	relations_deleteByQuery sSelect

END SUB


SUB relations_deleteCategories( nID )

	DIM	sSelect
	sSelect = "SELECT * " _
			&	"FROM tabcategorymap " _
			&	"WHERE " _
			&	" TabID = " & nID & ";"
	
	relations_deleteByQuery sSelect
	
END SUB


SUB relations_cleanup( sQuery, aArray(), sPrefix )
	DIM	oRS
	SEt oRS = dbQueryUpdate( g_DC, sQuery )
	
	IF NOT oRS.EOF THEN
	
		DIM	i
		DIM	bFound
		DIM	oRel
		SET oRel = oRS(sPrefix & "ID" )
	
		oRS.MoveFirst		
		DO UNTIL oRS.EOF
			bFound = FALSE
			FOR i = 0 TO UBOUND(aArray)
				IF 0 < aArray(i) THEN
					IF CLNG(oRel) = aArray(i) THEN
						bFound = TRUE
						aArray(i) = 0
						EXIT FOR
					END IF
				END IF
			NEXT 'i
			IF NOT bFound THEN
				oRS.Delete 1 'adAffectCurrent
			END IF
			oRS.MoveNext
		LOOP
	END IF
	
	Set oRS.ActiveConnection = Nothing
	oRS.Close
	SET oRS = Nothing
END SUB



SUB relations_cleanupCategories( n, aPreds() )

	DIM	sSelect
	sSelect = "SELECT * " _
			&	"FROM tabcategorymap " _
			&	"WHERE " _
			&	" TabID=" & n _
			&	";"

	relations_cleanup sSelect, aPreds, "Category"

END SUB


SUB relations_cleanupTabs( n, aSuccs() )

	DIM	sSelect
	sSelect = "SELECT * " _
			&	"FROM tabcategorymap " _
			&	"WHERE " _
			&	" CategoryID=" & n _
			&	";"

	relations_cleanup sSelect, aSuccs, "Tab"

END SUB


FUNCTION relations_needToAdd( aArray() )
	DIM	i
	FOR i = 0 TO UBOUND(aArray)
		IF 0 < aArray(i) THEN
			relations_needToAdd = TRUE
			EXIT FUNCTION
		END IF
	NEXT 'i
	relations_needToAdd = FALSE
END FUNCTION


FUNCTION relations_checkError( aArray() )
	DIM	i
	FOR i = 0 TO UBOUND(aArray)
		IF aArray(i) < 0 THEN
			relations_checkError = TRUE
			EXIT FUNCTION
		END IF
	NEXT 'i
	relations_checkError = FALSE
END FUNCTION



SUB relations_addCategories( n, aArray() )

	DIM	sSelect
	sSelect = "SELECT * " _
			&	"FROM tabcategorymap " _
			&	"WHERE " _
			&	" TabID=" & n _
			&	";"

	DIM g_RS
	SET g_RS = dbQueryUpdate( g_DC, sSelect )
	
	DIM	i
	
	FOR i = 0 TO UBOUND(aArray)
	
		IF 0 < aArray(i) THEN
		
			g_RS.AddNew
		
			g_RS.Fields("TabID") = n
			g_RS.Fields("CategoryID") = aArray(i)
			g_RS.Update
			
		END IF

	NEXT 'i

	SET g_RS.ActiveConnection = Nothing
	g_RS.Close
	
	SET g_RS = Nothing

END SUB


SUB relations_addTabs( n, aArray() )

	DIM	sSelect
	sSelect = "SELECT * " _
			&	"FROM tabcategorymap " _
			&	"WHERE " _
			&	" CategoryID = " & n & ";"

	DIM	oRS
	SET oRS = dbQueryUpdate( g_DC, sSelect )
	
	DIM	i
	
	FOR i = 0 TO UBOUND(aArray)
	
		IF 0 < aArray(i) THEN
		
			oRS.AddNew
		
			oRS.Fields("CategoryID") = n
			oRS.Fields("TabID") = aArray(i)
			oRS.Update
			
		END IF

	NEXT 'i

	SET oRS.ActiveConnection = Nothing
	oRS.Close
	
	SET oRS = Nothing

END SUB


FUNCTION relations_trim( sRelations )
	DIM	sList
	sList = TRIM(sRelations)
	sList = REPLACE( sList, " ", vbLF, 1, -1, vbTextCompare )
	sList = REPLACE( sList, vbCRLF, vbLF, 1, -1, vbTextCompare )
	sList = REPLACE( sList, vbCR, vbLF, 1, -1, vbTextCompare )
	DO UNTIL "" = sList
		IF vbLF = LEFT(sList,1) THEN
			sList = MID(sList,2)
		ELSEIF vbLF = RIGHT(sList,1) THEN
			sList = LEFT(sList,LEN(sList)-1)
		ELSEIF 0 < INSTR(sList,vbLF&vbLF) THEN
			sList = REPLACE( sList, vbLF&vbLF, vbLF, 1, -1, vbTextCompare )
		ELSE
			EXIT DO
		END IF
	LOOP
	relations_trim = sList
END FUNCTION


FUNCTION commitCategories( n, sCategories )

	commitCategories = TRUE
	
	DIM	sRelations
	sRelations = relations_trim( sCategories )
	IF "" = sRelations THEN
		relations_deleteCategories n
	ELSE

		DIM	aPreds()
	
		relations_getArrayFromString aPreds, sRelations
		
		IF relations_checkError( aPreds ) THEN
			commitCategories = FALSE
		ELSE
			relations_cleanupCategories n, aPreds
		
			IF relations_needToAdd( aPreds ) THEN
				relations_addCategories n, aPreds
			END IF
		END IF
	END IF

END FUNCTION


''' Successors = Tabs
''' Predecessors = Categories


FUNCTION commitTabs( n, sTabs )

	commitTabs = TRUE
	
	DIM	sRelations
	sRelations = relations_trim( sTabs )
	IF "" = sRelations THEN
		relations_deleteTabs n
	ELSE
		DIM	aSuccs()
	
		relations_getArrayFromString aSuccs, sRelations
		
		IF relations_checkError( aSuccs ) THEN
			commitTabs = FALSE
		ELSE
			relations_cleanupTabs n, aSuccs
		
			IF relations_needToAdd( aSuccs ) THEN
				relations_addTabs n, aSuccs
			END IF
		END IF
	END IF

END FUNCTION





%>