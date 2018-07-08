<!--#include file="index_tools.asp"-->
<%








SUB outputJavaScript
%>

<script type="text/javascript" language="JavaScript">
<!--
function replaceWindowURL( win, url )
{
	win.location.href = url;
}

function goPage( s )
{
	replaceWindowURL( window, s );
}

function getTimeString( oTime )
{
	var hours = oTime.getHours();
	var minutes = oTime.getMinutes();
	var meridian;
	
	// determine the meridian
	if ( 11 < hours )
	{
		// 12 o'clock will go to zero
		//  but it will be fixed below
		hours = hours - 12;
		meridian = "PM";
	}
	else
	{
		meridian = "AM";
	}
	
	// fix the parts for display
	if ( 0 == hours )
		hours = 12;
	if ( minutes < 10 )
		minutes = "0" + minutes;

	var timeString = hours + ":" + minutes + " " + meridian;
	
	return timeString;
}

function getDateString( dNow )
{
	var aMonths = new Array( "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" );
	var dd = dNow.getDate();
	var mm = dNow.getMonth();
	var yy = dNow.getFullYear();
	return aMonths[mm] + " " + dd + ", " + yy;
}



function getTodayString()
{
	var dNow = new Date();
	var sDOW = new Array( "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" );
	var nDOW = dNow.getDay();
	var sTime = getTimeString( dNow );
	var sDate = getDateString( dNow );
	return '<font face="sans-serif" size="1">&nbsp;' + sDOW[nDOW] + ', ' + sDate + ' ' + sTime + '</font>';
}

//-->
</script>
<script language="JavaScript1.1">
<!--
function replaceWindowURL( win, url )
{
	win.location.replace( url );
}
//-->
</script>
<script language="JavaScript1.2">
<!--

function showTime()
{
	// this is here to safeguard us from the wrong browsers
	if ( !document.all  &&  !document.layers )
		return;

	var clock = getTodayString();
	
	// here is the secret to changing a "span" or "div" dynamically.
	if ( document.layers )
	{
		var	oInnerDoc = document.layers["dynclock"].document;
		oInnerDoc.write( clock );
		oInnerDoc.close();
	}
	else if ( document.all )
	{
		dynclock.innerHTML = clock;
	}
	
	setTimeout( "showTime()", 15*1000 );
}

//if ( document.all )
//	setTimeout( "showTime()", 15*1000 );

//-->
</script>

<%
END SUB


SUB outputTitle
%>


<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <th valign="bottom" align="left">&nbsp;<font size="5"><%=g_sPageTitle%></font></th>
    <td valign="bottom" align="right">
<%
IF FALSE THEN
%>
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="bottom" align="right">
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td bgcolor="#006600"><img border="0" src="images/pie_brx.gif"
                  width="16" height="16"></td>
              </tr>
            </table>
          </td>
          <td valign="bottom" align="left">
            <table border="0" cellspacing="0" cellpadding="6">
              <tr>
                <th bgcolor="#006600" align="left"><font color="#FFFFFF"
                  size="5">Content Maintenance</font></th>
              </tr>
            </table>
          </td>
        </tr>
      </table>
<%
END IF
%><a target="_blank" href="../"><img border="0" src="theme/homeicon.gif"></a> &nbsp;&nbsp;
    </td>
  </tr>
</table>
<%
END SUB

SUB outputTimeBlock
%>
<table border="0" cellspacing="0" width="100%">
  <tr>
    <td align="right">
    <a href="../"><img border="0" src="theme/homeicon.gif"></a> &nbsp;&nbsp;</td>
  </tr>
</table>
<%
END SUB

FUNCTION genGoPage( sFile )
	genGoPage = "javascript:goPage('" & sFile & "')"
END FUNCTION
FUNCTION genLabel( bLink, sFile, sLabel )
	IF bLink THEN
		genLabel = "<a href=""" & sFile & """ class=""navigationtab"">" & Server.HTMLEncode(sLabel) & "</a>"
	ELSE
		genLabel = sLabel
	END IF
END FUNCTION

SUB outputTabs
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" height="6" class="navtabbox"><spacer type="block" height="1" width="1"></td>
  </tr>
  <tr>
    <td width="100%" align="right" class="navtabbox">
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <%
DIM i
DIM sBGColor

FOR i = 0 TO nFileCount-1
	IF i = g_nIndex THEN
		sBGColor = "#FFFFFF"
	ELSE
		sBGColor = "#CCCCCC"
	END IF
	aFileSplit = SPLIT( aFileList(i), vbTAB, -1, vbTextCompare )
          %>
          <td bgcolor="<%=sBGColor%>" valign="top"><img border="0"
            src="images/pie_tl_gray.gif" width="8" height="8"></td>
          <th bgcolor="<%=sBGColor%>"><%=genLabel(i <> g_nIndex, aFileSplit(kFI_Name), aFileSplit(kFI_Title))%></th>
          <td bgcolor="<%=sBGColor%>" valign="top" align="right"><img border="0"
            src="images/pie_tr_gray.gif" width="8" height="8"></td>
          <%
	IF i < nFileCount-1 THEN
          %>
          <td width="3"></td>
          <%
	END IF
NEXT 'i
          %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
END SUB

SUB outputPad
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" height="6"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<%
END SUB

outputJavaScript
outputTitle
'outputTimeBlock
outputTabs
outputPad

%>