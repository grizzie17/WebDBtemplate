<% 

' UpcomingEvents
'	mo,fr		indicates every monday and friday
'

'	Currently inherrited from default
'g_sUpcomingEventsSchedule = "mo,fr"

'g_UE_nDurationCalendar = 6

'g_UE_nDurationForum = 8

g_UE_nDurationAnnounce = 6

g_sDBUser = "grizzie_dbadmin"		'grizzie_dbadmin
g_sDBPassword = "170101Admin!"		'170101Admin!

g_sDBAdmin = "grizzie_dbadmin"
g_sDBAdminPassword = "170101Admin!"
g_bDBDSN = false
IF LCASE(Request.ServerVariables("SERVER_NAME")) = "localhost" THEN
	'g_sDBProvider = "SQLOLEDB"
	g_sDBProvider = "SQLNCLI11"
	g_sDBSource = "dbtemplate"
	g_sDBServer = "localhost"
	DIM	dbFilename
	dbFilename = findFileUpTree("app_data\" & g_sDBSource & ".mdf")
	'g_sDBConnectString = "" _
	'	&	"Provider=" & g_sDBProvider & ";" _
	'	&	"Driver={SQL Native Client};" _
	'	&	"Server=.\SQLEXPRESS;" _
	'	&	"AttachDbFilename=" & dbFilename & ";" _
	'	&	"Database=dblocal;" _
	'	&	"User Instance=True;" _
	'	&	"Connect Timeout=30;" _
	'	&	""
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
'g_sDistrictDomain = LCASE(Request.ServerVariables("SERVER_NAME"))


g_bCalendarPrefaceOption = FALSE




'g_sCalendarList = ""
'g_sCalendarHiddenList = ""

%>