<%
%>

<div class="noprint">
<%
SUB outputByLine
	DIM	sHost
	sHost = Request.ServerVariables("HTTP_HOST")
	Response.Write "<" & "script type=""text/javascript"" src=""http://"
	IF "localhost" = sHost THEN
		Response.Write "localhost/BearConsultingGroup"
	ELSE
		Response.Write "BearConsultingGroup.com"
	END IF
	Response.Write "/designby_small.js""></" & "script>" & vbCRLF
END SUB
outputByLine
%>
</div>
<%
%>
