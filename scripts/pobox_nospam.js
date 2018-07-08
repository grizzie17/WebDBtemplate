// nospam version


var g_at = '&#64;';
var g_sDom = 'Mid' + 'State' + 'Wings' + '&#46;' + 'org';
g_sDom = postoffice_getHostname();


function postoffice_getHostname()
{
	var xx = document.location.hostname;
	var i = xx.lastIndexOf(".");
	if ( 0 < i )
	{
		var j = xx.lastIndexOf(".",i-1);
		if ( 0 < j )
			xx = xx.substr(j+1);
	}
	return xx;
}


function postoffice()
{
	setTimeout( "timerPostoffice()", 0.1*1000 );
}


function timerPostoffice()
{
	var aSpans = document.getElementsByTagName( "SPAN" );
	var i;
	var j;
	var oSpan;
	var sContent;
	var sTemp;
	var sLink;
	for ( i = 0; i < aSpans.length; ++i )
	{
		oSpan = aSpans[i];
		switch ( oSpan.className )
		{
		case "pobox":
			sTemp = oSpan.innerHTML;
			sContent = new String(sTemp);
			j = sContent.indexOf( "[" );
			if ( 0 < j )
				sLink = poWithMessage( sContent );
			else
				sLink = poPlainAddress( sContent );
			if ( 0 < sLink.length )
			{
				oSpan.innerHTML = sLink;
				oSpan.className = "poboxsent";
			}
			break;
		case "dial":
			sTemp = oSpan.innerHTML;
			sContent = new String(sTemp);
			j = sContent.indexOf( "[" );
			if ( 0 < j )
				sLink = telWithMessage( sContent );
			else
				sLink = telPlainNumber( sContent );
			if ( 0 < sLink.length )
			{
				oSpan.innerHTML = sLink;
				oSpan.className = "dialed";
			}
			break;
		default:
			break;
		}
	}
}


function poWithMessage( sContent )
{
	var	i = sContent.indexOf( "[" );
	var sM;
	var sA;
	
	sM = sContent.substr(0,i);
	sM = poTrim( sM );
	sA = sContent.substr(i+1);
	i = sA.indexOf( "]" );
	if ( 0 < i )
		sA = sA.substr(0, i);
	sA = poParseAddress( sA );
	if ( 0 < sM.length )
		return sM;
	else
		return poBuildLink( sA, sA );
}

function telWithMessage( sContent )
{
	var	i = sContent.indexOf( "[" );
	var sM;
	var sA;
	
	sM = sContent.substr(0,i);
	sM = poTrim( sM );
	sA = sContent.substr(i+1);
	i = sA.indexOf( "]" );
	if ( 0 < i )
		sA = sA.substr(0, i);
	sA = telParseNumber( sA );
	if ( 0 < sM.length )
		return telBuildLink( sA, sM );
	else
		return telBuildLink( sA, sA );
}

function poPlainAddress( sContent )
{
	var sA = poParseAddress( sContent );
	if ( 0 < sA.length )
		return poBuildLink( sA, sA );
	else
		return "";
}

function telPlainNumber( sContent )
{
	var sA = telParseNumber( sContent );
	if ( 0 < sA.length )
		return telBuildLink( sA, sA );
	else
		return "";
}


function poLTrim( str )
{
	for( var k = 0; k < str.length && poIsWhitespace( str.charAt(k) ); k++)
		;
	return str.substring(k, str.length);
}

function poRTrim( str )
{
	for( var j=str.length-1; 0 <= j && poIsWhitespace(str.charAt(j)) ; j--)
		;
	return str.substring(0, j+1);
}

function poTrim( str )
{
	return poLTrim( poRTrim( str ));
}

var k_whitespaceChars = " \t\n\r\f";

function poIsWhitespace(charToCheck)
{
	return (k_whitespaceChars.indexOf(charToCheck) != -1);
}





function poBuildLink( sA, sT )
{
	if ( sA == sT )
		return " ";
	else
		return sT;
}

function telBuildLink( sA, sT )
{
	if ( sA == sT )
		return sA;
	else
		return sT + " (" + sA + ")";
}


function poParseAddress( sAddress )
{
	var sA;
	var sQ;
	var aa;
	var sT;
	var i;
	
	// trim white space
	sA = "" + poTrim( sAddress );
	
	// anything to the right of the v-bar will be the subject of the mail
	i = sA.indexOf( "|" );
	if ( 0 < i )
	{
		// no processing will be done on this
		sQ = sA.substr(i+1);
		sA = sA.substr(0,i);
		sA = poTrim( sA );
		sQ = poTrim( sQ );
	}
	else
	{
		sQ = "";
	}
	
	// the asterix is used to "hide" the at-sign
	i = sA.indexOf( "*" );
	if ( 0 < i )
	{
		do
		{
			sT = sA.substr(i+1);
			sA = sA.substr(0,i) + g_at + sT;
			i = sA.indexOf( "*" );
		} while ( 0 < i );
	}
	else
	{
		sA = sA + g_at + g_sDom;
	}
	
	// semicolons are used to "hide" periods
	i = sA.indexOf( ":" );
	if ( 0 < i )
	{
		do
		{
			sT = sA.substr(i+1);
			sA = sA.substr(0,i) + "&#46;" + sT;
			i = sA.indexOf( ":" );
		} while ( 0 < i );
	}
	
	// spaces are used to confuse things; just collapse out
	i = sA.indexOf( " " );
	if ( 0 < i )
	{
		do
		{
			sT = sA.substr(i+1);
			sA = sA.substr(0,i) + sT;
			i = sA.indexOf( " " );
		} while ( 0 < i );
	}
	
	if ( 0 < sQ.length )
	{
		sA = sA + "?Subject=" + escape(sQ);
	}
	
	return sA;
}


function telParseNumber( sAddress )
{
	var sA;
	var aa;
	var i;
	
	// trim white space
	sA = poTrim( sAddress );

	// the asterix is used to "hide" the dash
	i = sA.indexOf( "*" );
	if ( 0 < i )
	{
		do
		{
			sT = sA.substr(i+1);
			sA = sA.substr(0,i) + "-" + sT;
			i = sA.indexOf( "*" );
		} while ( 0 < i );
	}
	
	// spaces are used to confuse things; just collapse out
	i = sA.indexOf( " " );
	if ( 0 < i )
	{
		do
		{
			sT = sA.substr(i+1);
			sA = sA.substr(0,i) + sT;
			i = sA.indexOf( " " );
		} while ( 0 < i );
	}
	
	// pounds are used to "hide" extensions
	i = sA.indexOf( "#" );
	if ( 0 < i )
	{
		do
		{
			sT = sA.substr(i+1);
			sA = sA.substr(0,i) + " x" + sT;
			i = sA.indexOf( "#" );
		} while ( 0 < i );
	}
	
	return sA;
}



if (window.addEventListener)
	window.addEventListener("load", postoffice, false)
else if (window.attachEvent)
	window.attachEvent("onload", postoffice)
else
	setTimeout( "postoffice()", 2.0*1000 );


