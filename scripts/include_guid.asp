<%


FUNCTION makeGUID()
	makeGUID = ""
	DIM	oTypeLib
	
	SET oTypeLib = Server.CreateObject("Scriptlet.Typelib")
	IF NOT Nothing IS oTypeLib THEN
		DIM	sGUID
		
		sGUID = TRIM(oTypeLib.Guid)
		SET oTypeLib = Nothing
		sGUID = REPLACE(sGUID,"{","")
		sGUID = REPLACE(sGUID,"}","")
		makeGUID = TRIM(sGUID)
	END IF
END FUNCTION



FUNCTION getGuidChars(s)

	getGuidChars = "000000000000000000"

	DIM	oReg
	DIM	oMatchList

	SET oReg = NEW RegExp
	oReg.Pattern = "^(\d+|[A-Za-z]+|-){8,}"
	oReg.IgnoreCase = TRUE
	oReg.Global = TRUE
	SET oMatchList = oReg.Execute( s )
	IF 0 < oMatchList.Count THEN
		getGuidChars = LEFT(s, oMatchList(0).length)
	END IF
	

END FUNCTION



%>