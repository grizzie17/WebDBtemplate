



function replaceV( sContent, sSearch, sReplace )
{
	var	i,j;
	var	sA;
	var	sT;
	
	i = sContent.indexOf( sSearch );
	if ( -1 < i )
	{
		j = 0;
		sA = "";
		do
		{
			if ( j < i )
				sA += sContent.substring(j,i);
			sA += sReplace;
			j = i + sSearch.length;
			i = sContent.indexOf( sSearch, j );
		} while ( 0 < i );
		sA += sContent.substr(j);
		return sA;
	}
	else
	{
		return sContent;
	}
}


function fixupString( sString )
{
	var sContent = sString;
	
	if ( 0 < sContent.length )
	{
		var sNewLine = String.fromCharCode(0x0D,0x0A);
		sContent = replaceV( sContent, String.fromCharCode(0x85), "..." );
		sContent = replaceV( sContent, String.fromCharCode(0x2026), "..." );
		sContent = replaceV( sContent, String.fromCharCode(0x91), "'" );
		sContent = replaceV( sContent, String.fromCharCode(0x92), "'" );
		sContent = replaceV( sContent, String.fromCharCode(0xB4), "'" );	// acute
		sContent = replaceV( sContent, String.fromCharCode(0x2018), "'" );
		sContent = replaceV( sContent, String.fromCharCode(0x2019), "'" );
		sContent = replaceV( sContent, String.fromCharCode(0x93), '\"' );
		sContent = replaceV( sContent, String.fromCharCode(0x94), '\"' );
		sContent = replaceV( sContent, String.fromCharCode(0x201C), '\"' );
		sContent = replaceV( sContent, String.fromCharCode(0x201D), '\"' );
		sContent = replaceV( sContent, String.fromCharCode(0x96), "--" );
		sContent = replaceV( sContent, String.fromCharCode(0x97), "--" );
		sContent = replaceV( sContent, String.fromCharCode(0x2013), "--" );
		sContent = replaceV( sContent, String.fromCharCode(0x2014), "--" );
		sContent = replaceV( sContent, sNewLine+String.fromCharCode(0x95), sNewLine+"*" );	// bullet
		sContent = replaceV( sContent, sNewLine+String.fromCharCode(0xB7), sNewLine+"*" );	// middot
		sContent = replaceV( sContent, sNewLine+String.fromCharCode(0x2022), sNewLine+"*" );	// bullet
					
	}
	return sContent;
}


function fixupTextArea( sName )
{
	var sNewContent = "";
	
	var oThis = document.getElementById( sName );
	if ( oThis )
	{
		var sContent;
		
		if ( oThis.innerText )
			sContent = oThis.innerText;
		else
			sContent = oThis.innerHTML;
		
		if ( 0 < sContent.length )
		{
		
			sNewContent = fixupString( sContent );
			if ( oThis.innerText )
			{
				oThis.innerText = sNewContent;
			}
			else
			{
				sNewContent = replaceV( sNewContent, sNewLine, "<br>" );
				oThis.innerHTML = sNewContent;
			}
		}
	}
}


function fixupTextInput( sName )
{
	var sNewContent = "";
	
	var oThis = document.getElementById( sName );
	if ( oThis )
	{
		var sContent;
		
		sContent = oThis.value;
		
		if ( 0 < sContent.length )
		{
		
			sNewContent = fixupString( sContent );
			
			oThis.value = sNewContent;
		}
	}
}






