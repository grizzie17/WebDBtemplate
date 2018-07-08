<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Year-at-a-glance</title>
</head>

<body>

<form method="GET" action="calendar_yag.asp">
	<p>&nbsp;</p>
	<table border="0" id="table1" cellpadding="4" style="border-collapse: collapse">
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>Style</td>
		</tr>
		<tr>
			<td>Beginning Year:</td>
			<td><select size="1" name="year">
	<%
		DIM	y, i, j
		y = YEAR(NOW)
		i = y - 2
		FOR i = y - 1 TO y + 8
			Response.Write "<option "
			IF i = y THEN Response.Write "selected "
			Response.Write "value=""" & i & """>" & i & "</option>" & vbCRLF
		NEXT 'i
	%>
	</select></td>
			<td>Year Title:</td>
			<td><input type="text" name="styleYear" size="20"></td>
		</tr>
		<tr>
			<td>Beginning Month:</td>
			<td><select size="1" name="firstMonth">
	<option value="1" selected>January</option>
	<option value="2">February</option>
	<option value="3">March</option>
	<option value="4">April</option>
	<option value="5">May</option>
	<option value="6">June</option>
	<option value="7">July</option>
	<option value="8">August</option>
	<option value="9">September</option>
	<option value="10">October</option>
	<option value="11">November</option>
	<option value="12">December</option>
	</select></td>
			<td>Month Title:</td>
			<td><input type="text" name="styleMonthTitle" size="20"></td>
		</tr>
		<tr>
			<td>Start of Week:</td>
			<td><select size="1" name="firstWeekday">
	<option selected value="1">Sunday</option>
	<option value="2">Monday</option>
	<option value="3">Tuesday</option>
	<option value="4">Wednesday</option>
	<option value="5">Thursday</option>
	<option value="6">Friday</option>
	<option value="7">Saturday</option>
	<option value="-2">Monday w-e Stacked</option>
	<option value="-7">Stacked Saturday</option>
	</select>&nbsp;
			<input type="checkbox" name="daynameabbrev" value="ON" id="fp1" checked><label for="fp1">Abbrev.</label></td>
			<td>Day Title:</td>
			<td><input type="text" name="styleDayTitle" size="20"></td>
		</tr>
		<tr>
			<td>Weekend:</td>
			<td><select size="1" name="defWeekend">
			<option selected value="fl">First+Last</option>
			<option value="f2">First Two</option>
			<option value="l2">Last Two</option>
			</select></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>Page Layout:</td>
			<td><select size="1" name="layout">
			<option selected>Portrait</option>
			<option>Landscape</option>
			</select></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>Month Layout:</td>
			<td><select size="1" name="Progression">
			<option selected>Horizontal</option>
			<option>Vertical</option>
			</select></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
	</table>
	<p><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>

</body>

</html>