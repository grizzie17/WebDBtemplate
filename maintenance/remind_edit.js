


function onSelectMonth( oSelect )
{
	var	n;
	var	oForm;
	
	oForm = document.formDate;
	n = oSelect.selectedIndex;
	oForm.yearlyDayn_month.selectedIndex = n;
	oForm.yearlyWDay_month.selectedIndex = n;
	oForm.single_month.selectedIndex = n;
}

function onSelectHebrewMonth( oSelect )
{
	var	n;
	var	oForm;
	
	oForm = document.formDate;
	n = oSelect.selectedIndex;
	oForm.hebrewDayn_month.selectedIndex = n;
	oForm.hebrewWDay_month.selectedIndex = n;
	//oForm.single_month.selectedIndex = n;
}

function onWeeknumber( oSelect )
{
	var	n;
	var	oForm;
	
	oForm = document.formDate;
	n = oSelect.selectedIndex;
	oForm.yearlyWDay_weekNumber.selectedIndex = n;
	oForm.monthlyWDay_weekNumber.selectedIndex = n;
}

function onSelectWeekday( oSelect )
{
	var	n;
	var	oForm;
	
	oForm = document.formDate;
	n = oSelect.selectedIndex;
	oForm.monthlyWDay_weekday.selectedIndex = n;
	oForm.yearlyWDay_weekday.selectedIndex = n;
	oForm.hebrewWDay_weekday.selectedIndex = n;
}

function onSelectCycle( oSelect )
{
	var	n;
	var	oForm;
	
	oForm = document.formDate;
	n = oSelect.selectedIndex;
	oForm.monthlyWDay_cycle.selectedIndex = n;
	oForm.monthlyDayn_cycle.selectedIndex = n;
	onSelectMonthIndex( oForm.monthlyWDay_index );
}



function onSelectMonthIndex( oSelect )
{
	var	n;
	var i;
	var	oForm;
	
	oForm = document.formDate;
	n = oForm.monthlyWDay_cycle.selectedIndex;
	i = oForm.monthlyWDay_cycle.options[n].value - 1;
	n = oSelect.selectedIndex;
	if ( i < n )
		n = i;
	oForm.monthlyWDay_index.selectedIndex = n;
	oForm.monthlyDayn_index.selectedIndex = n;
}

function onDay( oInput )
{
	var	n;
	var	oForm;
	
	n = oInput.value;
	if ( isNaN(n) )
	{
		window.alert("Field requires numeric value");
		oInput.focus();
		oInput.click();
	}
	else
	{
		oForm = document.formDate;
		oForm.hebrewDayn_day.value = n;
		oForm.yearlyDayn_day.value = n;
		oForm.monthlyDayn_day.value = n;
		oForm.single_day.value = n;
	}
}

function selectItemOnValue( objOpt, sValue )
{
	var nMaxOptions;
	var	n;

	nMaxOptions = objOpt.length;
	for ( n = 0; n < nMaxOptions; ++n )
	{
		if ( sValue == objOpt.options[n].value )
		{
			objOpt.selectedIndex = n;
			return true;
		}
	}
	return false;
}


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


function fixupContent( oThis )
{
	//window.alert( "fixupContent" );
	//var oThis = document.getElementById( "Content" );
	if ( oThis )
	{
		var sContent;
		
		if ( oThis.innerText )
			sContent = oThis.innerText;
		else
			sContent = oThis.innerHTML;
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
						
			if ( oThis.innerText )
			{
				oThis.innerText = sContent;
			}
			else
			{
				sContent = replaceV( sContent, sNewLine, "<br>" );
				oThis.innerHTML = sContent;
			}
		}
	}
}




function submitValidate( oForm )
{
}



