<%
DIM	g_bCookiesEnabled
DIM	g_sCookiesTester
g_bCookiesEnabled = FALSE


SUB cookies_testEnabled()
	
	DIM	n
	
	RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
	n = ROUND( RND * 100)
	
	DIM	sTester
	sTester = g_sCookiePrefix & "_Cookie" & n
	g_sCookiesTester = sTester
	
	Response.Cookies(g_sCookiePrefix & "_CookiesTester") = sTester
	Response.Cookies(sTester) = CSTR(NOW)
	
END SUB

cookies_testEnabled




%>