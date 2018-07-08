<%

DIM g_dLocalPageDateModified
SUB LocalPageGetDate()
	DIM	oFile
	SET oFile = g_oFSO.GetFile( Request.ServerVariables("PATH_TRANSLATED") )
	DIM	dLastMod
	g_dLocalPageDateModified = oFile.DateLastModified
	SET oFile = Nothing
END SUB
LocalPageGetDate



%>
<div id="pageblock">

