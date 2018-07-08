<%

g_UE_nDurationAnnounce = 6

g_sDBUser = "grizzie_dbadmin"
g_sDBPassword = "170101Admin!"

g_sDBAdmin = "grizzie_dbadmin"
g_sDBAdminPassword = "170101Admin!"
g_bDBDSN = false
IF LCASE(Request.ServerVariables("SERVER_NAME")) = "localhost" THEN
	g_sDBProvider = "SQLNCLI11"
	g_sDBSource = "dbtemplate"
	g_sDBServer = "localhost"
	DIM	dbFilename
	'dbFilename = findFileUpTree("app_data\" & g_sDBSource & ".mdf")
    dbFilename = Server.MapPath("/app_data") & "\" & g_sDBSource & ".mdf"
	g_sDBConnectString = "" _
		&	"Provider=" & g_sDBProvider & ";" _
		&	"Server=(localdb)\MSSQLLocalDB;" _
		&	"AttachDbFilename=" & dbFilename & ";" _
		&	"Trusted_Connection=Yes;" _
		&	"Timeout=30;" _
		&	""
ELSE
	g_sDBSource = "dbtemplate"
END IF

g_sChapterNeighbors = "al-h,al-y,al-t"
g_sDistrictDomain = "Alabama-gwrra.org"


g_bCalendarPrefaceOption = FALSE

%>
