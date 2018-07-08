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
g_htmlFormat_downloadFunc = "pagePicture"

DIM sText
%>
<html>

<head>
    <title>Picture Help</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<style>
<!--
 /* Font Definitions */
p.htmlformatcaption, li.htmlformatcaption, div.htmlformatcaption
	{
	margin-right:0in;
	margin-left:0in;
	font-size:12.0pt;
	;
	color:#336699;}
p.htmlformatpullquote, li.htmlformatpullquote, div.htmlformatpullquote
	{
	margin-right:0in;
	margin-left:0in;
	background:#99CCFF;
	font-size:12.0pt;
	font-family:"Arial",sans-serif;
	font-style:italic;}
p.htmlformatsidebar, li.htmlformatsidebar, div.htmlformatsidebar
	{
	margin-right:0in;
	margin-left:0in;
	background:#99CCFF;
	font-size:10.0pt;
	font-family:"Arial",sans-serif;
	font-style:italic;}
p.htmlformat, li.htmlformat, div.htmlformat
	{
	margin-right:0in;
	margin-left:0in;
	font-size:12.0pt;
	;}
-->
</style>

</head>

<body lang=EN-US link=blue vlink=navy>
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
%>

<h1>Pictures</h1>

<p>Start by saving all of the pictures onto
your computer that you want displayed on your website.  **Hint: save all of
your pictures in a folder that you can navigate to easily from the upload
screen** </p>

<p>Go to http://dekalbhcl.org/maintenance
page, and log on using your credentials given to you by Elizabeth Griswold.  </p>

<p>Select the “Pages” tab from the list across the top of your page.&nbsp; </p>

<p><img width=624 height=93
id="Picture 8" src="htmlformathelp/maintenance.jpg"></p>

<p>You will find the pencil  <img style="width:16px" src="htmlformathelp/edit.gif"> under the word
<b>Edit</b> in the list of pages you have created.  Select the icon corresponding to
the page you want to add your picture to.  </p>

<p><br>
[ADD SCREENSHOT of dummy page with list to show where the pencil is]</p>

<p>Go to the text that you want to add the
picture to, and follow the below instructions on how to embed the picture into
the paragraphs.</p>

<p>[Daddy, please put your step by step
instructions here for embedding the picture]</p>

<p>[ADD SCREENSHOT of dummy page for adding pictures]</p>

<p>After embedding the pictures into the
text, you will scroll down and upload the photos to your page.  This area will
look like the below image. Label the picture USING THE SAME verbiage you used
in your embedded text.  Click on the Browse button and navigate to your picture
and select the picture.  You may also edit the size in this section as well,
480 is the recommended size.</p>

<p><img width=624 height=178
id="Picture 7" src="htmlformathelp/pictureupload.jpg"></p>

<p>-------------------</p>

<p>You can specify pictures that can be
embedded in your text.&nbsp;The picture is specified similarly to the text
styles using double braces (&quot;{{&quot;, &quot;}}&quot;).&nbsp; The easiest
method of specifying a picture is:</p>

<p>{{picture some-label}}</p>

<p>where the &quot;some-label&quot; is
the label that you provided when you uploaded the file to the web-server.</p>

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

<p>A caption may be specified by providing
text following a colon.&nbsp; Caption text may also be styled.&nbsp; Caption
text is not compatible with the &quot;width=percent%&quot; option.&nbsp;
Caption text should not be used with pictures that are associated with the
label of a web-address.</p>

<p>&nbsp;&nbsp;&nbsp; {{picture some-label : Caption text}}</p>

<p>The available optional arguments are
&quot;align&quot;, &quot;border&quot; and &quot;width&quot;.&nbsp; The
following is a description of each.</p>

<ul type=disc>
 <li>align=left -- the picture floats to
     the left margin and text wraps to the right</li>
 <li>align=right -- the picture floats to
     the right margin and text wraps to the left</li>
 <li>left -- same as align=left</li>
 <li>right -- same as align=right</li>
 <li>border=number -- specifies the
     number of pixels to be used for the border</li>
 <li>noborder -- specifies that no border
     should be used (same as border=0)</li>
 <li>width=number -- specifies the
     width of the picture in pixels</li>
 <li>width=percent% -- specifies
     the width of the picture as a percent of the paragraph width (ignored if a
     caption is specified)</li>
 <li>number -- same as
     width=number</li>
 <li>percent% -- same as
     width=percent%</li>
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

<p>You may also specify a picture by using
one of the terms: &quot;picture&quot;, &quot;pict&quot;, &quot;image&quot;,
&quot;img&quot;</p>

<p><b>Please reduce your pictures before
uploading them to the web.</b></p>


</body>

</html>
<%
    SET g_oFSO = NOTHING
%>
