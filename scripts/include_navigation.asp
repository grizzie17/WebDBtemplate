<%


PUBLIC aFileSplit
DIM g_sPage
DIM g_sPageTitle
DIM	g_sFile
DIM	g_nIndex
DIM	nLen
g_sPage = ""
g_sFile = ""
g_sPageTitle = ""




SUB outputTabs
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td align="right" valign="bottom">

<%

	
	Response.Write g_sTopNavTabs


%>
    </td>
  </tr>
</table>
<%
END SUB



SUB outputLogoBlock
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td id="headerPage" class="header">
<div id="headerImgGWRRA">
<a href="http://www.gwrra.org/" target="_blank"><img border="0" src="<%=g_sTheme%>gwrra_logo_100.gif" alt="Gold Wing Road Riders Association"></a>&nbsp; 
</div>
<div id="headerImgLogo">
<a href="./"><img border="0" src="<%=g_sTheme%>Logo-100.gif" alt="<%=g_sSiteName%>"></a> 
</div>
<div id="headerSiteTitle">
<%=g_sSiteTitle%>
</div>
<%
IF "" <> g_sSiteMotto THEN
%>
<div id="headerSiteMotto"><%=g_sSiteMotto%></div>
<%
END IF
%>
</td>
	</tr>
</table>
<%
END SUB
SUB outputTimeBlock
%>
<table border="0" cellspacing="0" width="100%">
  <tr>
    <td class="timeBlock">&nbsp; <%=g_sShortSiteName%> &nbsp;-&nbsp; <%=g_sSiteCity%> &nbsp; &nbsp;</td>
    <td class="timeBlock" align="right"><a target="_blank" href="http://www.alabama-gwrra.org/">Alabama District</a> &nbsp;-&nbsp; <a target="_blank" href="http://www.gwrra-regiona.org/">Region-A</a> &nbsp;-&nbsp; <a target="_blank" href="http://www.gwrra.org/">Gold Wing Road Riders Association</a> &nbsp;</td>
  </tr>
</table>
<script type="text/javascript" language="javascript" src="scripts/calcNextMeeting.js"></script>
<!--table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" bgcolor="#000000" height="2"><spacer type="block" height="1" width="1"></td>
  </tr>
</table-->
<%
END SUB






SUB outputPad2
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" height="6"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<%
END SUB




outputLogoBlock

outputTabs
outputTimeBlock
'outputPad
IF Response.Buffer THEN Response.Flush

%>