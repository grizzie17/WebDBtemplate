
function validateRequired( oForm )
{

	var	bValid = true;
	var aFields;
	var nFields;
	var	oField;
	var	i;
	var idx;
	

	//oForm = document.EditForm;
	if ( oForm.getElementsByTagName )
		aFields = oForm.getElementsByTagName("SELECT");
	else
		aFields = oForm.all.tags("SELECT");
	nFields = aFields.length;
	
	for ( i = 0 ; i < nFields ; i++ )
	{
		oField = aFields[i];
		idx = oField.selectedIndex;
		if ( 0 == idx )
		{
			if ( -1 == oField.options[idx].value )
				oField.selectedIndex = -1;
		}
		if ( "required" == oField.className )
		{
			if ( -1 == oField.selectedIndex )
			{
				bValid = false;
				//oField.focus();
				break;
			}
		}
	}
	if ( bValid )
	{
		if ( oForm.getElementsByTagName )
			aFields = oForm.getElementsByTagName("INPUT");
		else
			aFields = oForm.all.tags("INPUT");
		nFields = aFields.length;
		for ( i = 0 ; i < nFields ; i++ )
		{
			oField = aFields[i];
			if ( "required" == oField.className )
			{
				if ( "" == oField.value )
				{
					bValid = false;
					break;
				}
			}
		}
	}
	if ( ! bValid )
	{
		if ( oField )
		{
			alert( "Required fields are missing" );
			oField.focus();
			oField.click();

			bValid = false;
		}
	}
	return bValid;

}

