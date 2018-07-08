<%
OPTION EXPLICIT

DIM	g_oFSO
SET g_oFSO = CreateObject("Scripting.FileSystemObject")
%>
<!--#include file="htmlformat.asp"-->
<%

DIM	sPicturePath

sPicturePath = Server.MapPath("htmlformathelp")



FUNCTION pagePicture( sLabel )
	DIM	sFile
	IF 1 = INSTR(1,sLabel,"http://",vbTextCompare) THEN
		pagePicture = sLabel
	ELSE
		sFile = sLabel & ".gif"
		IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( sPicturePath, sFile )) THEN
			sFile = sLabel & ".jpg"
			IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( sPicturePath, sFile )) THEN
				sFile = sLabel & ".png"
				IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( sPicturePath, sFile )) THEN sFile = ""
			END IF
		END IF
		pagePicture = "htmlformathelp/" & sFile
	END IF
END FUNCTION

g_htmlFormat_pictureFunc = "pagePicture"


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Simple Text - Help</title>
<style type="text/css">
<!--

html, body
{
	width: 400px;
}

h1, h2, h3, h4, h5, h6
{
	font-family: Arial,Helvetica,Verdana,sans-serif;
}

h3.htmlformat
{
	border-bottom: 3px solid black;
}

h4.htmlformat
{
	border-bottom: 2px solid black
}

h5.htmlformat
{
	border-bottom: 1px solid #666666
}

h6.htmlformat
{
	border-bottom: 1px solid #999999
}


hr.htmlformat
{
	color: #3399CC;
}

.htmlformatcaption
{
	color: #336699;
	font-size:	small;
}

.htmlformatpullquote
{
	background-color: #99CCFF;
	font-style: italic;
	font-family: Arial,Helvetica,Verdana,sans-serif;
}

.htmlformatsidebar
{
	background-color: #99CCFF;
	font-style: italic;
	font-size:	x-small;
	font-family: Arial,Helvetica,Verdana,sans-serif;
}


-->
</style>
</head>

<body bgcolor="#FFFFFF" vlink="#000080">
<%
DIM	nTA
nTA = 0
SUB outputRaw( sText, nRows )
	DIM	sTemp
	
	sTemp = REPLACE( Server.HTMLEncode(sText), vbLF, "<br>" & vbCRLF, 1, -1, vbTextCompare )
	sTemp = REPLACE( sTemp, "  ", "&nbsp; ", 1, -1, vbTextCompare )

	Response.Write "<code>"
	Response.Write sTemp
	Response.WRite "</code>"
END SUB

SUB outputExample( sText )
	Response.Write "<p>Example:</p>" & vbCRLF
	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""2"">" & vbCRLF
	Response.Write "<tr>"
	Response.Write "<td width=""100%"" bgcolor=""#FFFFCC"">"
	outputRaw sText, 1
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>" & vbCRLF

	Response.Write "<p>This text would produce:</p>" & vbCRLF

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""2"">" & vbCRLF
	Response.Write "<tr>"
	Response.Write "<td width=""100%"" bgcolor=""#CCFFFF"">"
    Response.Write htmlFormat( sText )
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>" & vbCRLF
END SUB

SUB outputNoFormat( sText )
	'Response.Write "<p>Example with no formatting:</p>" & vbCRLF
	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""2"">" & vbCRLF
	Response.Write "<tr>"
	Response.Write "<td width=""100%"" bgcolor=""#FFCCFF"">"
    Response.Write Server.HTMLEncode( sText )
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>" & vbCRLF
END SUB
%>
<div style="width: 350px">
<h2><a name="top"></a>Simple Text Formatting</h2>
Simple Text Formatting (STF) is designed to provide a method to create very 
readable text when viewed as plain-text and when processed with the 
html-formatter very good looking and readable web-content.&nbsp; Most of STF can 
be done without the use of any kind of meta-tags.&nbsp; Some of the more 
advanced features do require some very simple meta-tags but they are not as 
complex as html and provide even richer content.&nbsp; The primary features of STF
include:
<ul>
  <li><a href="#para">Simple Paragraphs</a></li>
  <li><a href="#lists">Lists</li></a>
  <ul>
  <li><a href="#bullet">Bullet Lists</a></li>
  <li><a href="#numbers">Numbered Lists</a></li>
  <li><a href="#outline">Outline Lists</a></li>
  <li><a href="#indentedblocks">Indented Blocks</a></li>
</ul>
  
  <li><a href="#header">Headers</a></li>
  <li><a href="#rule">Horizontal Rules (Lines)</a></li>
  <li><a href="#link">Web-Site Address Links</a></li>
</ul>
Advanced Features:
<ul>
  <li><a href="#styles">Text Styles</a></li>
	<li><a href="#fonts">Fonts</a></li>
  <li><a href="#pictures">Pictures</a></li>
  <li><a href="#pullquotes">Pull Quotes</a></li>
</ul>
<hr>
<h3><a name="para"></a>Simple Paragraphs</h3>
<p>The simplest form of STF is creating paragraphs.&nbsp; Paragraphs are 
created by simply providing an empty line between blocks of text.</p>
<%

DIM	sText
sText = _
		"This is a block of text with no embedded returns.  " _
		&	"In other words when the user entered the text they " _
		&	"did not press the ""enter"" key while typing this paragraph." & vbLF _
		&	vbLF _
		&	"Here is another block of text, again with no embedded returns." & vbLF _
		&	vbLF _
		&	"And yet another block of text."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>If an embedded return is found (user pressed &quot;enter&quot; without
producing a blank line) the text following the return will start on a new line
but a new paragraph will not be produced.</p>
<%

sText =	"This is a paragraph with many lines." & vbLF _
		&	"Second line of paragraph" & vbLF _
		&	"Another line after pressing ""enter""" & vbLF _
		&	vbLF _
		&	"This is another paragraph" & vbLF _
		&	"Yet another line" & vbLF _
		&	"Short line again"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>The following is an example of what the preceding plain-text would be 
formatted as by a web-browser if the STF routines were not used to process the 
text.</p>
<!--webbot bot="PurpleText" preview="Output no format" -->
<%
	outputNoFormat sText
%>
<h4>Centered Paragraphs</h4>
<p>You can also create centered paragraphs by prefixing the first line of a
paragraph with a circumflex (&quot;^&quot;).</p>
<%

sText =	"^ This is a paragraph with many lines that will be centered." & vbLF _
		&	"Second line of paragraph" & vbLF _
		&	"Another line after pressing ""enter""" & vbLF _
		&	vbLF _
		&	"^ This is another paragraph" & vbLF _
		&	"Yet another line" & vbLF _
		&	"Short line again"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>The circumflex prefix as used for centered-paragraphs is a good example of 
how many of the remaining features of STF are specified.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="lists"></a>Lists</h3>
<p>Lists are special forms of paragraphs.&nbsp; The entire list must be 
delimited by empty lines.&nbsp; Each list-item is a text line that begins with a 
special character or sequence of characters.&nbsp; Each subsequent line (following a return) 
that begins with the special character or sequence is considered the next 
list-item.</p>
<h4><a name="bullet"></a>Bullet Lists</h4>
<p> Bullet lists are created by beginning a line with an asterisk (&quot;*&quot;).&nbsp; 
Each subsequent line (following a return) that begins with an asterisk is 
considered the next point of the bullet list.</p>
<%
sText = _
		"* This is the first bulleted point." & vbLF _
		&	"If you press ""enter"" (without a blank line) the following text " _
		&	"will start on a new line but still be part of the same bullet point." & vbLF _
		&	"* This is the second bullet point." & vbLF _
		&	"* This is the third bullet."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>It is important to note that bullet lists (as well as all other lists) must 
have an empty line before the first list-item. Notice in the following example 
the formatted result is not treated as a bullet list but is simply treated as a single paragraph.</p>
<%
sText = _
		"The following is a bullet list:" & vbLF _
		&	"* This is the first bulleted point." & vbLF _
		&	"* This is the second bullet point." & vbLF _
		&	"* This is the third bullet."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>By adding an empty line before the list the formatting routines can convert 
the asterisk list into a true bullet list.</p>
<%
sText = _
		"The following is a bullet list:" & vbLF _
		&	vbLF _
		&	"* This is the first bulleted point." & vbLF _
		&	"* This is the second bullet point." & vbLF _
		&	"* This is the third bullet."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>

<p>Bullet lists can also create nested (indented) lists by specifying more than
one asterisk at the beginning of the line.&nbsp; Each asterisk represents
another level of nesting.&nbsp; As each list is nested the
bullet shape changes. Starting with the first level the bullets are; solid round disc, open circle,
solid square, then the cycle repeats.</p>
<%
sText = _
			"* This is the first bulleted point." & vbLF _
		&	"* This is the second bullet point." & vbLF _
		&	"** This is a nested bullet point" & vbLF _
		&	"** Second nested point" & vbLF _
		&	"*** Third level nested point" & vbLF _
		&	"*** Third level nested point two" & vbLF _
		&	"* This is the third bullet."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>A special method of representing a bullet list can be specified by starting 
each bullet point with a single dash (&quot;-&quot;).&nbsp; There must not be more than a 
single dash and there must be at least two dash-bullet items.&nbsp; Nested 
(indented) dash-bullet items are not supported.</p>
<%
sText = _
		"-This is the first bulleted point." & vbLF _
		&	"If you press ""enter"" (without a blank line) the following text " _
		&	"will start on a new line but still be part of the same bullet point." & vbLF _
		&	"-This is the second bullet point." & vbLF _
		&	"-This is the third bullet."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Bullet lists are delimited by empty lines.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h4><a name="numbers"></a>Numbered Lists</h4>
<p>Numbered lists are nearly identical to bullet lists with the exception that 
each line begins with a number sign (&quot;#&quot;).</p>
<%
sText = _
			"# This is the first numbered point." & vbLF _
		&	"Line after a return." & vbLF _
		&	"# This is the second number point." & vbLF _
		&	"# This is the third number."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>It is important to note that for &quot;all&quot; numbered lists there must be at least 
two numbered list-items. Notice in the following example that each poorly formed 
numbered item is in a separate paragraph.</p>
<%
sText = "" _
		&	"# This is a single numbered Point" & vbLF & vbLF _
		&	"# This is another single numbered point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>If you want a blank line separating points of a list separate the lines with 
 
a single vertical-bar (&quot;|&quot;).</p>
<%
sText = "" _
		&	"# This is a single numbered Point" & vbLF _
		&	"|" & vbLF _
		&	"# This is another single numbered point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Nested or Indented lists are also supported.&nbsp; As each list is nested the
numbering system changes. The numbering systems are; Arabic
numerals, lowercase letters,  lowercase
Roman numerals, and then remaining nesting is
identical to the bullet lists.</p>
<%
sText =	"# This is the first numbered point." & vbLF _
		&	"# This is the second number point." & vbLF _
		&	"## Nested numbered point" & vbLF _
		&	"## Second nested numbered point" & vbLF _
		&	"### Third nesting one" & vbLF _
		&	"### Third nesting two" & vbLF _
		&	"# This is the third number."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>In addition to  identifying a numbered list with number signs you can create 
a numbered list with numbers.&nbsp; Each number must be delimited with either a 
period ('.') or a close parenthesis (')').</p>
<%
sText =	"1) This is the first numbered point." & vbLF _
		&	"2. This is the second number point." & vbLF _
		&	"3.5 is a number that should not be treated as a numbered point." & vbLF _
		&	"  a. Nested numbered point" & vbLF _
		&	"  b. Second nested numbered point" & vbLF _
		&	"3 This is the third number." & vbLF _
		&	"4. This is the fourth number." & vbLF _
		&	"A) Nested numbered point" & vbLF _
		&	"B) Second nested numbered point" & vbLF _
		&	"5.) The fifth numbered point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>It is important to note that numbered lists created in this manner can only 
have one level of nesting.</p>
<p>It is important to note that for &quot;all&quot; numbered lists there must be at least 
two numbered list-items. Notice in the following example that each poorly formed 
numbered item is in a separate paragraph.</p>
<%
sText = "" _
		&	"1. This is a single numbered Point" & vbLF & vbLF _
		&	"2) This is another single numbered point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>The following example is another example of a &quot;numbered&quot; list that creates an 
&quot;alpha-numbered&quot; list.&nbsp; No nesting is currently supported for this form of 
list.&nbsp; Also please note the 4th item in the source and its final output.</p>
<%
sText =	"a) This is the first numbered point." & vbLF _
		&	"b. This is the second number point." & vbLF _
		&	"c) Third numbered point" & vbLF _
		&	"c) Fourth numbered point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Here is yet another example of a poorly formed numbered list.</p>
<%
sText = "" _
		&	"a) poorly formed alpha-list point" & vbLF & vbLF _
		&	"b. second poorly formed point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Numbered lists are delimited by empty lines.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h4><a name="outline"></a>Outline Lists</h4>
<p>Outline lists are nearly identical to numbered lists with the exception that 
each line begins with a percent sign (&quot;%&quot;) and they follow the numbering rules 
of outlines.</p>
<%
sText =	"% This is the first outline point." & vbLF _
		&	"Line after a return." & vbLF _
		&	"% This is the second outline point." & vbLF _
		&	"% This is the third outline point."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Here is yet another example of a poorly formed numbered list.</p>
<%
sText = "" _
		&	"% poorly formed outline-list point" & vbLF & vbLF _
		&	"% second poorly formed point"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Nested or Indented lists are also supported. As each list is nested the
numbering system changes according to the rules of
outlines.&nbsp; The numbering systems are; capital Roman numerals, capital-alpha
letters, Arabic
numerals, lowercase-alpha letters, lowercase Roman numerals, and then remaining nesting is
identical to the bullet lists.</p>
<%
sText =	"% This is the first outline point." & vbLF _
		&	"% This is the second outline point." & vbLF _
		&	"%% Nested outline point" & vbLF _
		&	"%% Second nested outline point" & vbLF _
		&	"%%% Yet another nested point" & vbLF _
		&	"%%%% Additional nested point" & vbLF _
		&	"%%%% Yep, another nested point" & vbLF _
		&	"% This is the third outline point."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Outline lists are delimited by empty lines.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h4><a name="indentedblocks"></a>Indented Blocks</h4>
<p>Indented Blocks are similar in nature to the lists but they differ in the 
fact that there are no bullets or numbers.&nbsp; They just support indenting as 
a block of text.&nbsp; Each paragraph that should be indented should be prefixed 
with a close (right) square bracket (&quot;]&quot;).</p>
<%
sText =	"] This is a block or paragraph of text that should be indented. " _
		&	"This can behave like any normal paragraph." & vbLF _
		&	"Line after a return."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Nested indented blocks are also supported.</p>
<%
sText =	"] This is the first block of the indented text." & vbLF _
		&	"This is the second line of the text." & vbLF _
		&	"]] Nested indented block" & vbLF _
		&	"Second line of nested block" & vbLF _
		&	"]]] Yet another nested block of text" & vbLF _
		&	"]]]] Additional nested text" & vbLF _
		&	"Yep, another line of text in the nested block" & vbLF _
		&	"] This is some text that is back at the original nesting level."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Indented blocks are delimited by empty lines.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="header"></a>Headers</h3>
<p>Headers are created by prefixing a line with an exclamation point
(&quot;!&quot;).&nbsp; Please note that just like paragraphs and lists, a header 
must be delimited with preceding and following empty lines.</p>
<%
sText =	"! This is a sample header" & vbLF _
		&	"with two lines of text" & vbLF _
		&	vbLF _
		&	"!! This is a smaller header" & vbLF _
		&	vbLF _
		&	"!!! This is an even smaller header" & vbLF _
		&	vbLF _
		&	"!!!! This is the smallest header"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Browsers that honor web-styles will render the headers with an extended rule (line) directly
under the header text.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="tables"></a>Tables</h3>
<p>Simple tables can be created by making regular blocks of data separated by vertical bars 
(&quot;|&quot;).&nbsp; Care needs to be taken that there are the same number of vertical bars on each row.</p>
<%
sText =	"" _
		&	"|A1|B1|C1|D1|" & vbLF _
		&	"|A2|B2|C2|D2|" & vbLF _
		&	"|A3|B3|C3|D3|" & vbLF _
		&	"|A4|B4|C4|D4|"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>It depends on how the styles have been set as to whether or not the table actually displays 
its borders or not.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="rule"></a>Horizontal Rule (line)</h3>
<p>A Horizontal Rule is specified by providing a line of dashes.&nbsp;&nbsp; If 
the line contains any characters other than dashes then the line does not 
generate a Horizontal Rule.&nbsp; It does not matter how many dashes there are 
on the line as long as they are all dashes.&nbsp; A heavy horizontal rule can be 
created by specifying a line of equal signs.</p>
<%
sText =	"Here is some simple paragraph text that is to be followed by a horizontal rule." & vbLF _
		&	vbLF _
		&	"----------" & vbLF _
		&	vbLF _
		&	"The dashes above should produce a horizontal rule." & vbLF _
		&	vbLF _
		&	"=========" & vbLF _
		&	vbLF _
		&	"The equal signs above should produce a heavy (or thick) horizontal rule." & vbLF _
		&	vbLF _
		&	"----- -- -----" & vbLF _
		&	vbLF _
		&	"The previous line should not produce a horizontal rule (notice the spaces). " &vbLF _
		&	"----------------" & vbLF _
		&	"The previous dashes will produce a rule even though it is embedded in a paragraph."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Browsers that honor web-styles will render the horizontal rules in a solid
color that complements the background color.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="link"></a>Web-Site Address Links</h3>
<p>The text is checked for web-site addresses.&nbsp; Any text
that starts with either &quot;http://&quot; or &quot;www.&quot; will be examined
and turned into a link to the identified site.&nbsp; E-Mail addresses are also
changed into &quot;mailto:&quot; links.</p>
<%
sText =	"You really should visit http://GrizzlyWeb.com. It is an excellent web portal." & vbLF _
		&	vbLF _
		&	"www.starwars.com is another really good site." & vbLF _
		&	vbLF _
		&	"Or just specify the web server, such as GrizzlyWeb.com or Sun-Sentinel.com, " _
		&	"but the domain should have a well known suffix unlike some-wild-server.zzz." & vbLF _
		&	"NT style servers are handled as well \\some-server.com\mount-point\somefile.doc." & vbLF _
		&	vbLF _
		&	"Please send email to Support@GrizzlyWeb.com if you have any questions."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>In addition to creating links from web-site and e-mail addresses you
can define an alternate label for the address by providing text
following the address enclosed in square brackets
(&quot;[&quot;, &quot;]&quot;).&nbsp; There must be exactly one space
between the end of the address and the open square bracket.</p>
<%
sText =	"You really should visit http://GrizzlyWeb.com [The Grizzly Web Links]. It is an excellent web portal." & vbLF _
		&	vbLF _
		&	"www.starwars.com [Star Wars] is another really good site." & vbLF _
		&	vbLF _
		&	"Please send email to Support@GrizzlyWeb.com [Bear Consulting Group Support] if you have any questions."
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>The list of supported protocols (which includes some extensions) and what they result in are 
as follows:</p>
<ul>
	<li>http: or https: -- results in a link that will open a new web browser to the specified 
	address</li>
	<li>file: -- opens a file browser to the specified address</li>
	<li>mailto: -- opens your mail program and creates a new message that is addressed to send a 
	message to the specified email address</li>
	<li>news: -- opens your news reader to the news directory specified</li>
	<li>address: -- opens a web browser showing a map of the street address specified</li>
</ul>
<p align="right"><a href="#top">Back to Top</a></p>
<hr>
<h2>Advanced Features</h2>
<p>The advanced formatting features of STF use some basic meta-tag like syntax.&nbsp; 
In general the syntax is:</p>
<blockquote>
<p>{{tagname options <b>:</b> text }}</p>
</blockquote>
<p>Where:</p>
<ul>
  <li>tagname - is the operation, such as a text style</li>
  <li>options - affects how &quot;tagname&quot; works</li>
  <li>text - what the &quot;tagname&quot; is affecting</li>
</ul>
<h3><a name="styles"></a>Text Styles</h3>
<p>You can also specify some more advanced text styles.&nbsp; The entered text
is not quite as obvious but it still is readable and understandable for what you
are trying to accomplish.&nbsp; The text styles include bold, italic, underline,
small, and large.&nbsp; You specify the styled text by surrounding the
text with double braces (&quot;{{&quot;, &quot;}}&quot;) and prefixing the
styled text with a format tag followed by a colon (&quot;:&quot;)</p>
<%
sText =	"This is some {{b:Bolded text}}." & vbLF _
		&	"This is some {{i:Italicized text}}." & vbLF _
		&	"This is some {{u:Underlined text}}." & vbLF _
		&	"And some {{large:Large size text}}" & vbLF _
		&	"And some {{small:Small size text}}" & vbLF _
		&	"Some {{strike:Strike through text}}" & vbLF _
		&	"And text with superscript X{{sup:2}}" & vbLF _
		&	"And a subscript X{{sub:3}}"


%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>It is also possible to use the text styles in combination with each other by
nesting the double braces.</p>
<%
sText =	"This text is {{style b i:Bolded and Italicized}}." & vbLF _
		&	"{{b:Bold, {{i:Italicized, {{u:Underlined and {{big:Big}}}}}}}}" & vbLF
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Text styles also support specifying colors.</p>
<%
sText = "{{b:" _
		&	"{{darkred:darkred}}, " _
		&	"{{maroon:maroon}}, " _
		&	"{{brown:brown}}, " _
		&	"{{chocolate:chocolate}}, " _
		&	"{{red:red}}, " _
		&	"{{coral:coral}}, " _
		&	"{{sandybrown:sandybrown}}, " _
		&	"{{darkorange:darkorange}}, " _
		&	"{{orange:orange}}, " _
		&	"{{goldenrod:goldenrod}}, " _
		&	"{{gold:gold}}, " _
		&	"{{yellow:yellow}}, " _
		&	"{{darkkhaki:darkkhaki}}, " _
		&	"{{darkgoldenrod:darkgoldenrod}}, " _
		&	"{{yellowgreen:yellowgreen}}, " _
		&	"{{lime:lime}}, " _
		&	"{{limegreen:limegreen}}, " _
		&	"{{green:green}}, " _
		&	"{{seagreen:seagreen}}, " _
		&	"{{darkseagreen:darkseagreen}}, " _
		&	"{{darkgreen:darkgreen}}, " _
		&	"{{olive:olive}}, " _
		&	"{{mediumaquamarine:mediumaquamarine}}, " _
		&	"{{darkcyan:darkcyan}}, " _
		&	"{{teal:teal}}, " _
		&	"{{aqua:aqua}}, " _
		&	"{{cyan:cyan}}, " _
		&	"{{turquoise:turquoise}}, " _
		&	"{{mediumturquoise:mediumturquoise}}, " _
		&	"{{darkturquoise:darkturquoise}}, " _
		&	"{{cadetblue:cadetblue}}, " _
		&	"{{steelblue:steelblue}}, " _
		&	"{{skyblue:skyblue}}, " _
		&	"{{cornflowerblue:cornflowerblue}}, " _
		&	"{{deepskyblue:deepskyblue}}, " _
		&	"{{blue:blue}}, " _
		&	"{{mediumblue:mediumblue}}, " _
		&	"{{darkblue:darkblue}}, " _
		&	"{{navy:navy}}, " _
		&	"{{indigo:indigo}}, " _
		&	"{{purple:purple}}, " _
		&	"{{pink:pink}}, " _
		&	"{{violet:violet}}, " _
		&	"{{orchid:orchid}}, " _
		&	"{{fuchsia:(fuchsia)}}, " _
		&	"{{magenta:magenta}}, " _
		& vbLF _
		&	"{{silver:silver}}, " _
		&	"{{darkgray:darkgray}}, " _
		&	"{{gray:gray}}, " _
		&	"{{dimgray:dimgray}}, " _
		&	"{{black:black}} " _
		&	"}}"
%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Text styles can be specified with several different names:</p>
<ul>
  <li>&quot;b&quot;, &quot;bold&quot;</li>
  <li>&quot;i&quot;, &quot;italic&quot;</li>
  <li>&quot;u&quot;, &quot;underline&quot;, &quot;_&quot;</li>
  <li>&quot;big&quot;, &quot;large&quot;</li>
  <li>&quot;small&quot;</li>
	<li>&quot;strike&quot;, &quot;s&quot;</li>
	<li>&quot;sup&quot;, &quot;^&quot;</li>
	<li>&quot;sub&quot;</li>
	<li>&quot;nowrap&quot;</li>
  <li>&quot;style&quot;</li>
</ul>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="fonts"></a>Fonts</h3>
<p>You can also specify font or font-family names in a similar fashion to styles.&nbsp; The 
fonts can include any of the generic-font-names: 
serif, sans-serif, cursive, fantasy, or monospace.&nbsp; </p>
<p>The method of specifying a pull-quote is:</p>
<p>&nbsp;&nbsp;&nbsp; {{font <i>font-name</i><b>:</b> <i>text to apply 
font</i>}}</p>
<p>In addition to the generic-font-names you may specify any font name.&nbsp; It 
is important to note that if you need to specify a font whose name contains 
spaces you must replace the spaces with underscores (&quot;_&quot;).&nbsp; As an example 
if you were specifying &quot;Cooper Black&quot; the name must be specified as &quot;Cooper_Black&quot;.</p>
<%
sText =	"{{font serif:Some serif text}}" & vbLF _
		&	"{{font sans-serif:and some sans-serif text}}" & vbLF _
		&	"{{font cursive:then some cursive text}}" & vbLF _
		&	"{{font fantasy:and neat Fantasy text}}" & vbLF _
		&	"{{font monospace:Monospaced text}}" & vbLF _
		&	"{{font arial_black:Arial Black text}}" & vbLF _
		&	"{{font cooper_black:Cooper Black text}}" & vbLF


%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="pictures"></a>Pictures</h3>
<p>You can specify pictures that can be embedded in your text.&nbsp;The picture is
specified similarly to the text styles using double braces (&quot;{{&quot;,
&quot;}}&quot;).&nbsp; The easiest method of specifying a picture is:</p>
<blockquote>
<p>{{picture <i>some-label</i>}}</p>
</blockquote>
<p>where the &quot;<i>some-label</i>&quot; is the label that you provided when you 
uploaded the file to the web-server.</p>
<%
sText =	"Some image {{picture web-paint}} that is embedded in the line." & vbLF _
		&	vbLF _
		&	"{{picture bear-paw right}} Picture with alignment that will word wrap etc. " _
		&	"Using the align option makes it real easy to embed your pictures into your text." & vbLF _
		&	vbLF _
		&	"Little picture {{picture bear-paw 20}}" & vbLF _
		&	vbLF _
		&	"Picture used with web-address Support@GrizzlyWeb.com [{{picture email-envelope}}]"

%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>A caption may be specified by providing text following a colon.&nbsp; Caption text may also be
styled.&nbsp; Caption text is not compatible with the &quot;width=percent%&quot;
option.&nbsp; Caption text should not be used with pictures that are associated
with the label of a web-address.</p>
<p>&nbsp;&nbsp;&nbsp; {{picture some-label <b>:</b> <i>Caption text</i>}}</p>
<p>The available optional arguments are &quot;align&quot;, &quot;border&quot; and
&quot;width&quot;.&nbsp; The following is a description of each.</p>
<ul>
  <li>align=left -- the picture floats to the left margin and text wraps to the
    right</li>
  <li>align=right -- the picture floats to the right margin and text wraps to
    the left</li>
  <li>left -- same as align=left</li>
  <li>right -- same as align=right</li>
  <li>border=<i>number</i> -- specifies the number of pixels to be used for the
    border</li>
  <li>noborder -- specifies that no border should be used (same as border=0)</li>
  <li>width=<i>number</i> -- specifies the width of the picture in pixels</li>
  <li>width=<i>percent%</i> -- specifies the width of the picture as a percent
    of the paragraph width (ignored if a caption is specified)</li>
  <li><i>number</i> -- same as width=<i>number</i></li>
  <li><i>percent%</i> -- same as width=<i>percent%</i></li>
</ul>
<%
sText =	"{{picture dancing-frog align=left:Dancing {{i:Frog}} Cartoon}} " _
		&	"Picture with alignment and caption.  Take note that the word {{i:""Frog""}} is {{i:italicized}} in the caption. " _
		&	"The align option works very well with the caption option." & vbLF _
		&	vbLF _
		&	"Picture with no border Support@GrizzlyWeb.com [{{picture mailbox noborder}}] used as a web-address"

%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>

<p>You may also specify a picture by using one of the terms:
&quot;picture&quot;, &quot;pict&quot;, &quot;image&quot;, &quot;img&quot;</p>

<p><b>Please reduce your pictures before uploading them to the web.</b></p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="pullquotes"></a>Pull Quotes</h3>
<p>Pull Quotes are used to highlight or emphasize key phrases within your
text.&nbsp;  The method of specifying a pull-quote is:</p>
<p>&nbsp;&nbsp;&nbsp; {{pullquote<b>:</b> <i>text to be quoted</i>}}</p>
<p>You can also abbreviate pullquote as simply &quot;q&quot;.</p>
<p>&nbsp;&nbsp;&nbsp; {{q<b>:</b> <i>text to be quoted</i>}}</p>
<%
sText =	"This is some normal text with a {{pullquote:key phrase}} sprinkled in here and there." & vbLF _
		&	vbLF _
		&	"In general you should avoid specifying more than one pull-quote per paragraph. " _
		&	"Limit quotes to {{q left:significant facts or important sound-bites}} that people will remember. " _
		&	"After all it doesn't make sense to provide a paragraph and then virtually quote its entire content."

%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p> You can also logically un-quote a section of your quoted text.&nbsp; This allows
you to skip a portion of the quoted text to emphasize certain sections.&nbsp;
The unquoted section is represented by an ellipses (&quot;...&quot;) in the
pull-quote.&nbsp; Un-quote is specified by either using the tag &quot;unquote&quot;,
&quot;uq&quot;, or &quot;...&quot;.</p>
<%
sText =	"In general you should avoid specifying more than {{q 150:one pull-quote per paragraph{{...:. " _
		&	"Try to}} limit quotes to significant facts or important sound-bites}} that people will remember. " _
		&	"After all it doesn't make sense to provide a paragraph and then virtually quote its entire content."

%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Pull quotes support the width and align optional arguments that operate
identically to the corresponding picture arguments.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<h3><a name="sidebar"></a>Sidebars</h3>
<p>Sidebars are used to provide additional information in an article or text.&nbsp; 
Sidebars are similar in operation to pull-quotes, but the block of identified 
text is not replicated in the main paragraph.&nbsp; The text only shows up in 
the sidebar block.&nbsp;  The method of specifying a sidebar is:</p>
<p>&nbsp;&nbsp;&nbsp; {{sidebar<b>:</b> <i>text to be quoted</i>}}</p>
<p>You can also specify a sidebar as &quot;block&quot;.</p>
<p>&nbsp;&nbsp;&nbsp; {{block<b>:</b> <i>text to be quoted</i>}}</p>
<%
sText =	"This is some normal text with a {{sidebar:key phrase}} sprinkled in here and there." & vbLF _
		&	vbLF _
		&	"In general you should avoid specifying more than one pull-quote per paragraph. " _
		&	"Limit quotes to {{block left:significant facts or important sound-bites}} that people will remember. " _
		&	"After all it doesn't make sense to provide a paragraph and then virtually quote its entire content."

%>
<!--webbot bot="PurpleText" preview="Output Example" -->
<%
	outputExample sText
%>
<p>Sidebars support the width and align optional arguments that operate
identically to the corresponding picture arguments.</p>
<p align="right"><a href="#top">Back to Top</a></p>
<p><font color="#999999" size="2">This page uses the Simple Text Formatting
routines to display the contents of the light blue boxes.</font></p>
</div>
</body>

</html>