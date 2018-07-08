
function getElementByName( sName )
{
	var oInputs = document.getElementsByName( sName );
	if ( oInputs )
	{
		var	oInput = oInputs.item(0);
		return oInput;
	}
	else
	{
		return oInputs;
	}
}


function formMailValidateOnSubmit( oForm )
{
	var oRequired = getElementByName('mailRequired');
	var sRequired = oRequired.value;
	var aRequired = sRequired.split(",");
	var	bValid = true;
	
	for (var i = 0; i < aRequired.length; ++i )
	{
		oRequired = getElementByName(aRequired[i]);
		if ( oRequired )
		{
			if ( 0 == oRequired.value.length )
			{
				bValid = false;
				oRequired.className = "missing";
			}
			else
			{
				oRequired.className = "found";
			}
		}
	}
	return bValid;
}



function setMailSessionGUID()
{
	var o = getElementByName( "mailSessionGUID" );
	if ( o )
		o.value = g_sSessionGUID;
}


if (window.attachEvent)
	window.attachEvent("onload", setMailSessionGUID);
else if (window.addEventListener)
	window.addEventListener("load", setMailSessionGUID, false);
else
	setTimeout( "setMailSessionGUID()", 2.0*1000 );


/*

<form method="POST" action="../formmailsubmit.asp" onsubmit="return formMailValidateOnSubmit( this )">
	<input type="hidden" name="mailSessionGUID" value="">
	<input type="hidden" name="mailPageTitle" value="Submit Rider Education Information Survey">
	<input type="hidden" name="mailSubject" value="RocketCityWings.org - Rider Ed Info Survey">
	<input type="hidden" name="mailFromAddr" value="not if i cat i ons">
	<input type="hidden" name="mailTo" value="ed u ca to r">
	<input type="hidden" name="mailBCC" value="we b ma st er">
	<input type="hidden" name="mailOrderFields" value="FromName,mailReplyTo,phone,CourseERC,CourseARC,CourseTRC,CourseTTRC,CourseTC,CourseSRC,CoursePLP,CourseCoRider,CourseRoadCaptain,CourseGroupRiding,CourseMatureRider,CourseCrashSceneResponse,CourseFirstAid,CourseCPR,Comments">
	<input type="hidden" name="mailRequired" value="FromName,mailReplyTo,phone">
	<input type="text" name="mailReplyTo" size="35"><br>
mailPriority
mailRedirect
mailSendSuccess

*/