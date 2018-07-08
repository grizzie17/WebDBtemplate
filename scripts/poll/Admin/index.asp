<%Response.Buffer=TRUE%>
<%IF Request.Cookies("polllog") <> "xx" THEN Response.Redirect "login.asp"%>
<html>
<head>
<title>PacPoll V 4.0</title>
</head>
<body>
<h1>PacPoll V 4.0</h1>
<h3>Change Username and password!</h3>
<form method="POST" action="change.asp">
  <b><font size="4">Username</font></b><br>
  <input type="text" name="uid" size="20"><br>
  <b><font size="4">Password</font></b><br>
  <input type="text" name="pwd" size="20"><input type="submit" value="Submit" name="B1"></p>
</form>
<h3>To add a new poll, enter the question and four possible responses below. Then
select the color for the bar graph.</h3>
<form method="POST" action="adpoll.asp">
  <input type="hidden" name="T6" value=""><p><b>Enter
  the QUESTION for the poll. (required)</b><br>
  <input type="text" name="T1" size="84"><br>
  <b>Enter answer one. (required)</b>
  <br>
  <input type="text" name="T2" size="84"><br>
  <b>Enter answer two. (required)</b>
  <br>
  <input type="text" name="T3" size="84"><br>
  <b>Enter answer three. (optional)</b>
  <br>
  <input type="text" name="T4" size="84"><br>
  <b>Enter answer four. (optional)</b>
  <br>
  <input type="text" name="T5" size="84"><br>
  <b>Select Font Type</b><br>
  <select size="1" name="D1">
    <option selected value="Arial">Arial</option>
    <option value="Arial Narrow">Arial Narrow</option>
    <option value="Book Antiqua">Book Antiqua</option>
    <option value="Bookman Old Style">Bookman Old Style</option>
    <option value="Century Gothic">Century Gothic</option>
    <option value="Comic Sans MS">Comic Sans MS</option>
    <option value="Courier">Courier</option>
    <option value="Courier New">Courier New</option>
    <option value="Fixedsys">Fixedsys</option>
    <option value="Franklin Gothic Medium">Franklin Gothic Medium</option>
    <option value="Garamond">Garamond</option>
    <option value="Georgia">Georgia</option>
    <option value="Haettenschweiler">Haettenschweiler</option>
    <option value="Impact">Impact</option>
    <option value="Lucida Console">Lucida Console</option>
    <option value="Microsoft Sans Serif">Microsoft Sans Serif</option>
    <option value="Modern">Modern</option>
    <option value="Monotype Corsiva">Monotype Corsiva</option>
    <option value="MS Sans Serif">MS Sans Serif</option>
    <option value="Sans Serif">Sans Serif</option>
    <option value="MV Boli">MV Boli</option>
    <option value="Palatino Linotype">Palatino Linotype</option>
    <option value="Roman">Roman</option>
    <option value="Script">Script</option>
    <option value="Small Fonts">Small Fonts</option>
    <option value="System">System</option>
    <option value="Tahoma">Tahoma</option>
    <option value="Terminal">Terminal</option>
    <option value="Times New Roman">Times New Roman</option>
    <option value="Verdana">Verdana</option>
  </select><br>
  <b>Select Font Size</b><br>
  <select size="1" name="D2">
    <option selected value="1">8pt</option>
    <option value="2">10pt</option>
    <option value="3">12pt</option>
    <option value="4">14pt</option>
  </select><br>
  <b>Select Font Color</b><br>
  <select size="1" name="D3" style="font-weight: bold">
    <option selected value="FFFFFF">White</option>
    <option value="C0C0C0" style="color: #C0C0C0">#C0C0C0</option>
    <option value="808080" style="color: #808080">#808080</option>
    <option value="000000" style="color: #000000">#000000</option>
    <option value="000080" style="color: #000080">#000080</option>
    <option value="0000FF" style="color: #0000FF">#0000FF</option>
    <option value="3366CC" style="color: #3366CC">#3366CC</option>
    <option value="99CCFF" style="color: #99CCFF">#99CCFF</option>
    <option value="00FFFF" style="color: #00FFFF">#00FFFF</option>
    <option value="800000" style="color: #800000">#800000</option>
    <option value="993333" style="color: #993333">#993333</option>
    <option value="FF0000" style="color: #FF0000">#FF0000</option>
    <option value="FF0066" style="color: #FF0066">#FF0066</option>
    <option value="FF5050" style="color: #FF5050">#FF5050</option>
    <option value="FF6600" style="color: #FF6600">#FF6600</option>
    <option value="FF9966" style="color: #FF9966">#FF9966</option>
    <option value="FFFF00" style="color: #FFFF00">#FFFF00</option>
    <option value="660066" style="color: #660066">#660066</option>
    <option value="800080" style="color: #800080">#800080</option>
    <option value="9900CC" style="color: #9900CC">#9900CC</option>
    <option value="FF00FF" style="color: #FF00FF">#FF00FF</option>
    <option value="FF99FF" style="color: #FF99FF">#FF99FF</option>
    <option value="003300" style="color: #003300">#003300</option>
    <option value="333333" style="color: #333333">#333333</option>
    <option value="008000" style="color: #008000">#008000</option>
    <option value="808000" style="color: #808000">#808000</option>
    <option value="00FF00" style="color: #00FF00">#00FF00</option>
    <option value="669999" style="color: #669999">#669999</option>
    <option value="008080" style="color: #008080">#008080</option>
    <option value="996600" style="color: #996600">#996600</option>
  </select><br>
  <b>Select the color for the bar graph.<br>
  <select size="1" name="D4" style="font-weight: bold">
    <option selected value="FFFFFF">White</option>
    <option value="C0C0C0" style="color: #C0C0C0">#C0C0C0</option>
    <option value="808080" style="color: #808080">#808080</option>
    <option value="000000" style="color: #000000">#000000</option>
    <option value="000080" style="color: #000080">#000080</option>
    <option value="0000FF" style="color: #0000FF">#0000FF</option>
    <option value="3366CC" style="color: #3366CC">#3366CC</option>
    <option value="99CCFF" style="color: #99CCFF">#99CCFF</option>
    <option value="00FFFF" style="color: #00FFFF">#00FFFF</option>
    <option value="800000" style="color: #800000">#800000</option>
    <option value="993333" style="color: #993333">#993333</option>
    <option value="FF0000" style="color: #FF0000">#FF0000</option>
    <option value="FF0066" style="color: #FF0066">#FF0066</option>
    <option value="FF5050" style="color: #FF5050">#FF5050</option>
    <option value="FF6600" style="color: #FF6600">#FF6600</option>
    <option value="FF9966" style="color: #FF9966">#FF9966</option>
    <option value="FFFF00" style="color: #FFFF00">#FFFF00</option>
    <option value="660066" style="color: #660066">#660066</option>
    <option value="800080" style="color: #800080">#800080</option>
    <option value="9900CC" style="color: #9900CC">#9900CC</option>
    <option value="FF00FF" style="color: #FF00FF">#FF00FF</option>
    <option value="FF99FF" style="color: #FF99FF">#FF99FF</option>
    <option value="003300" style="color: #003300">#003300</option>
    <option value="333333" style="color: #333333">#333333</option>
    <option value="008000" style="color: #008000">#008000</option>
    <option value="808000" style="color: #808000">#808000</option>
    <option value="00FF00" style="color: #00FF00">#00FF00</option>
    <option value="669999" style="color: #669999">#669999</option>
    <option value="008080" style="color: #008080">#008080</option>
    <option value="996600" style="color: #996600">#996600</option>
  </select><br>
  Set the color for the poll background.<br>
  <select size="1" name="D5" style="font-weight: bold">
    <option selected value="FFFFFF">White</option>
    <option value="C0C0C0" style="color: #C0C0C0">#C0C0C0</option>
    <option value="808080" style="color: #808080">#808080</option>
    <option value="000000" style="color: #000000">#000000</option>
    <option value="000080" style="color: #000080">#000080</option>
    <option value="0000FF" style="color: #0000FF">#0000FF</option>
    <option value="3366CC" style="color: #3366CC">#3366CC</option>
    <option value="99CCFF" style="color: #99CCFF">#99CCFF</option>
    <option value="00FFFF" style="color: #00FFFF">#00FFFF</option>
    <option value="800000" style="color: #800000">#800000</option>
    <option value="993333" style="color: #993333">#993333</option>
    <option value="FF0000" style="color: #FF0000">#FF0000</option>
    <option value="FF0066" style="color: #FF0066">#FF0066</option>
    <option value="FF5050" style="color: #FF5050">#FF5050</option>
    <option value="FF6600" style="color: #FF6600">#FF6600</option>
    <option value="FF9966" style="color: #FF9966">#FF9966</option>
    <option value="FFFF00" style="color: #FFFF00">#FFFF00</option>
    <option value="660066" style="color: #660066">#660066</option>
    <option value="800080" style="color: #800080">#800080</option>
    <option value="9900CC" style="color: #9900CC">#9900CC</option>
    <option value="FF00FF" style="color: #FF00FF">#FF00FF</option>
    <option value="FF99FF" style="color: #FF99FF">#FF99FF</option>
    <option value="003300" style="color: #003300">#003300</option>
    <option value="333333" style="color: #333333">#333333</option>
    <option value="008000" style="color: #008000">#008000</option>
    <option value="808000" style="color: #808000">#808000</option>
    <option value="00FF00" style="color: #00FF00">#00FF00</option>
    <option value="669999" style="color: #669999">#669999</option>
    <option value="008080" style="color: #008080">#008080</option>
    <option value="996600" style="color: #996600">#996600</option>
  </select><br>
  Set the width of the poll in pixels.<br>
  <select size="1" name="D6">
  <option selected value="125">125</option>
  <option value="150">150</option>
  <option value="175">175</option>
  </select><br>
  <input type="submit" value="Submit" name="B1"></p>
</form>
<p>&nbsp;</p>
</b>
</body>
</html>