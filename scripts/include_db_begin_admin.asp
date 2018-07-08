<!--#include file="include_dbmssql.asp"-->
<%




DIM g_oFSO
SET g_oFSO = Server.CreateObject( "Scripting.FileSystemObject" )




DIM g_DC
SET g_DC = dbConnectAuth(g_sDBSource, g_sDBAdmin, g_sDBAdminPassword)




%>