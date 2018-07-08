<%


SUB makeWeatherStyles()
%>

<style type="text/css">

.weatherCurrentTemp
{
	position: relative;
	font-size: 24px;
	font-family: sans-serif;
}

.weatherCurrentText
{
	position: relative;
	font-size: 14px;
	font-family: sans-serif;
}


.weatherForecastDay
{
	position: relative;
	display: inline-block;
	font-size: 18px;
	font-family: sans-serif;
	min-width: 2em;
}

.weatherForecastText
{
	position: relative;
	font-size: 10px;
	font-family: sans-serif;
}

.weatherForecastHigh
{
	position: relative;
	color: red;
	font-size: 16px;
	font-family: sans-serif;
}

.weatherForecastLow
{
	position: relative;
	color: blue;
	font-size: 12px;
	font-family: sans-serif;
}




</style>

<%
END SUB


FUNCTION genweather( sZip )

	genweather = ""
	
	
	DIM	oXML
	
	SET oXML = rss_getXMLCached( "YahooWeather" & sZip & ".xml", _
					"http://xml.weather.yahoo.com/forecastrss?p=" & sZip , "", _
					"h", 1, "d" )
'	SET oXML = rss_getXML( "YahooWeather" & sZip & ".xml", "http://xml.weather.yahoo.com/forecastrss?p=" & sZip , "" )
	IF NOT Nothing IS oXML THEN

		DIM oXSL
		SET oXSL = xmldom_loadFile( Server.MapPath( "scripts/rssweather.xslt" ) )
		IF NOT oXSL IS Nothing THEN

			genweather = oXML.transformNode(oXSL)
			genweather = REPLACE( genweather, "_$$", "&" )

			SET oXSL = Nothing
		END IF
		SET oXML = Nothing

	END IF
END FUNCTION



FUNCTION rssweather( sZip )

	DIM	sData
	sData = ""
	DIM	sCacheFile
	sCacheFile = "weather" & sZip & ".htm"
	DIM	oFile
	SET oFile = cache_openTextFile( "site", sCacheFile, "h", 1, "d" )
	IF NOT oFile IS Nothing THEN
		sData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
	ELSE
		sData = genweather( sZip )
		IF "" <> sData THEN
			SET oFile = cache_makeFile( "site", sCacheFile )
			IF NOT oFile IS Nothing THEN
				oFile.Write sData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF
	END IF
	rssweather = sData
	
END FUNCTION


%>