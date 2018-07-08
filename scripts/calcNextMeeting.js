
function dateFromWeekNumber( nWeek, nWDay, nMonth, nYear )
{
	var dFirst = new Date( nYear, nMonth, 1 );
	var nFirstDay = dFirst.getDay();
	var x = 1 + ( nWDay - nFirstDay + 7 ) % 7;
	x += 7 * nWeek;
	
	return new Date( nYear, nMonth, x );
}


var g_dNowNextMeeting;

function decodeMeetingCalendar( s )
{
	var sMeeting = "Unknown";
	var sContent = new String( s );
	var	i = sContent.indexOf( " " );
	var nWeek;
	var nWDay;
	
	nWeek = new Number(sContent.substr(0,i)) - 1;
	nWDay = new Number(sContent.substr(i+1)) - 1;
	
	var dNow = g_dNowNextMeeting;
	var	nDay = dNow.getDate();
	var nMonth = dNow.getMonth();
	var nYear = dNow.getFullYear();

	var dNew = dateFromWeekNumber( nWeek, nWDay, nMonth, nYear );
	if ( dNew < dNow )
	{
		++nMonth;
		if ( 12 < nMonth )
		{
			++nYear;
			nMonth = 1;
		}
		dNew = dateFromWeekNumber( nWeek, nWDay, nMonth, nYear );
	}
	
	if ( dNew.toDateString() == dNow.toDateString() )
		sMeeting = "TODAY";
	else
		sMeeting = dNew.toLocaleDateString()

	return sMeeting;
}


function getMeetingCalendars()
{
	var aSpans = document.getElementsByTagName( "SPAN" );
	var i;
	var j;
	var oSpan;
	var sContent;
	var sLink;
	
	var dNow = new Date();
	var	nDay = dNow.getDate();
	var nMonth = dNow.getMonth();
	var nYear = dNow.getFullYear();
	g_dNowNextMeeting = new Date( nYear, nMonth, nDay );

	for ( i = 0; i < aSpans.length; ++i )
	{
		oSpan = aSpans[i];
		switch ( oSpan.className )
		{
		case "meetcalendar":
			sContent = oSpan.innerHTML;
			sLink = decodeMeetingCalendar( sContent );
			if ( 0 < sLink.length )
			{
				oSpan.innerHTML = sLink;
				oSpan.className = "meetingcalendar";
			}
			break;
		default:
			break;
		}
	}
}


if (window.attachEvent)
	window.attachEvent("onload", getMeetingCalendars);
else if (window.addEventListener)
	window.addEventListener("load", getMeetingCalendars, false);
else
	setTimeout( "getMeetingCalendars()", 2.0*1000 );


