<%@ Language=VBScript %>
<%
OPTION EXPLICIT

DIM	q_q

q_q = Request("q")


Response.Redirect "http://maps.live.com/default.aspx?v=2&tilt=-90&alt=-1000&where1=" & Server.URLEncode(q_q)
'Response.Redirect "http://maps.google.com/maps?q=" & q_q & "&iwloc=A&hl=en"




%>