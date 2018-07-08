<%@	Language=VBScript
	EnableSessionState=True %>
<%
Option Explicit
Response.Expires = 0
Response.Buffer = True


%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\include_logon.asp"-->
<!--#include file="scripts\include_db_begin.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<%


IF NOT authHasAccess( "CFGe" ) AND NOT authHasAccess( "SRVe" ) THEN
	Response.Redirect "./"
END IF


DIM	sID
DIM	sCategory
DIM	sBrand
sID = Request("page")
sCategory = Request("category")




%><html>

<head>
<title>Edit Site Configuration</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="robots" content="noindex">
<link rel="stylesheet" href="theme/style.css" type="text/css">
<script language="JavaScript" type="text/javascript" src="scripts/formvalidate.js"></script>
<script language="JavaScript" type="text/javascript" src="scripts/pageedit.js"></script>
<script type="text/javascript" language="JavaScript">
<!--

function doCancel()
{
	window.location.href = "./";
}

function validateForm()
{
	var oForm = document.EditForm;
	var sts;
	
	sts = validateRequired( oForm );
	//if ( sts )
	//	sts = checkRequiredPictureLabel( oForm );
	return sts;
}


function checkRequiredPictureLabel( theForm )
{
    var bMissing = false;
    var i;
    var oField;
    var sValue;
    
    i = 0;
    while ( i++ < 5 )
    {
    	oField = eval( theForm.name+".LabelPictureFile" + i );
       	sValue = oField.value;
    	oField = eval( theForm.name+".Label" + i );
    	if ( ! verifyLabel( oField ) )
    		bMissing = true;
    	if ( 0 < sValue.length )
    	{
    		if ( oField.value.length < 1 )
    			bMissing = true;
    	}
    	else
    	{
    		if ( 0 < oField.value.length )
    			bMissing = true;
    	}
    }
            
    if ( bMissing )
    {
        alert( "Uploading Pictures requires\nboth a label and the picture.");
        
        // false causes the form submission to be canceled
        return false;
    }
    else
    {
        return true;
    }
}




function verifyLabel( oLabel )
{
	var sValue = oLabel.value;
	if ( 0 < sValue.length )
	{
		var re = /^[\w-]+$/gi;
		var ar;
		ar = sValue.match( re );
		if ( null == ar )
		{
			alert( "Picture label contains invalid characters.");
			oLabel.focus();
			return false;
		}
		else
		{
			return true;
		}
	}
	else
	{
		return true;
	}
}



//-->
</script>
</head>

<body bgcolor="#FFFFFF" id="pageEdit">

<h1>Edit Site Configuration</h1>
<%

FUNCTION checkTrue( o )
	checkTrue = ""
	IF NOT ISNULL( o ) THEN
		IF CBOOL( o ) THEN
			checkTrue = "checked=""checked"""
		ELSE
			checkTrue = ""
		END IF
	END IF
END FUNCTION

FUNCTION makeSelect( x, y )
	IF UCASE(x) = UCASE(y) THEN
		makeSelect = " selected "
	ELSE
		makeSelect = " "
	END IF
END FUNCTION

FUNCTION MIN( a, b )
	IF a < b THEN
		MIN = a
	ELSE
		MIN = b
	END IF
END FUNCTION


FUNCTION makeSize( oType, oMaximum, nSize )
	SELECT CASE recNumber(oType)
	CASE 0	'string
		IF ISNULL(oMaximum) THEN
			makeSize = nSize
		ELSE
			makeSize = MIN(nSize, oMaximum.Value)
		END IF
	CASE 1	'integer
		makeSize = 10
	CASE ELSE
		makeSize = nSize
	END SELECT
END FUNCTION

FUNCTION makeMax( oType, oMaximum, oData )
	SELECT CASE recNumber(oType)
	CASE 0	'string
		IF ISNULL(oMaximum) THEN
			makeMax = oData.DefinedSize
		ELSE
			makeMax = MIN(oData.DefinedSize, oMaximum.Value)
		END IF
	CASE 1	'integer
		makeMax = 11
	CASE ELSE
		makeMax = oData.DefinedSize
	END SELECT
END FUNCTION





	DIM	sSelect
	sSelect = "" _
			&	"SELECT " _
			&		"* " _
			&	"FROM " & dbQ("config") & " " _
			&	"ORDER BY " & dbQ("Description") & " " _
			&	";"
	
	DIM	oRS
	DIM	RS
	SET RS = dbQueryRead( g_DC, sSelect )



	IF RS.EOF Then
		RS.Close
		Response.Write "<p>Requested Config Data not found</p>"
	ELSE
		DIM	oRID
		DIM	oName
		DIM	oDescription
		DIM	oAccess
		DIM	oType
		DIM	oData
		DIM	oMaximum

		SET	oName = RS.Fields("Name")
		SET	oDescription = RS.Fields("Description")
		SET	oAccess = RS.Fields("Access")
		SET	oType = RS.Fields("Type")
		SET	oMaximum = RS.Fields("Maximum")
		SET	oData = RS.Fields("DataValue")

		DIM	sName

		


%>
<form method="POST" action="configeditsubmit.asp" name="EditForm" onsubmit="return validateForm()">
	<div align="center">
		<table border="0" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<%
		RS.MoveFirst
		DO UNTIL RS.EOF
			IF authHasAccess( recString( oAccess ) ) THEN
				sName = recString( oName )
%>
			<tr>
				<td align="right" valign="top"><%=recString(oDescription) %></td>
				<td>
<%
				SELECT CASE oType.Value
				CASE 0,1
%>
				<input type="text" id="<%=sName %>" name="<%=sName %>" size="<%=makeSize(oType,oMaximum,40) %>" maxlength="<%=makeMax(oType,oMaximum,oData)%>" value="<%=Server.HTMLEncode(recString(oData))%>"/>
<%
					IF 0 = oType.Value THEN
%>
					<a href="javascript:fixupTextInput('<%=sName %>')">fix-special-chars</a>
<%
					END IF
				CASE 2
%>
					<input type="checkbox" id="<%=sName %>" name="<%=sName %>" <%=checkTrue(oData) %> value="true" />
<%
				END SELECT
%>
				<!--div class="instructions">
					Used for web-page titles</!--div-->
				</td>
			</tr>
<%
			END IF
			RS.MoveNext
		LOOP
%>
		</table>
		<p><input type="submit" value="Save" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="button" value="Cancel" name="B2" onclick="doCancel()"></p>
	</div>
</form>
<%

		RS.Close


	END IF
SET	RS = Nothing
%> 
<!--#include file="scripts\byline.asp"-->

</body>

</html>
<%
%><!--#include file="scripts\include_db_end.asp"-->