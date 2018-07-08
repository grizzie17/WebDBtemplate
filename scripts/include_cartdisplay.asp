<%

SUB cartDisplay()
%>
<div style="position: absolute; top:5px; right: 10px; z-index:100;">
<table border="0" cellspacing="0" cellpadding="1" align="right">
  <tr>
    <td align="right" class="cartcount">
    <%
    	DIM	sSelect
    	DIM	oRS
    	DIM	nCount
    	nCount = 0
    	
    	DIM	sQuote
    	sQuote = "'"
    	
    	sSelect = "" _
    		&	"SELECT " _
    		&		"SUM(Qty) AS CartCount " _
    		&	"FROM " _
    		&		"Orders " _
    		&	"WHERE " _
    		&		"Cookie = " & sQuote & sCartCookie & sQuote _
    		&	";"
    	
    	'Response.Write "<p>" & sSelect & "</p>" & vbCRLF
    	'Response.Flush
    	
    	SET oRS = dbQueryRead( g_DC, sSelect )
    	IF NOT oRS.EOF THEN
    		nCount = recNumber(oRS.Fields("CartCount"))
    	ELSE
    		nCount = 0
    	END IF
    	oRS.Close
    	SET oRS = Nothing
    	
    	IF nCount < 1 THEN
    		Response.Write "No Items in Cart"
    	ELSE
	    	Response.Write "<a href=""cartview.asp""  class=""cartcount""><b>"
%>
    <img border="0" src="images/icn_cart.gif" width="13" height="15">&nbsp;
<%
	    	IF 1 = nCount THEN
		    	Response.Write nCount & " Item in Cart"
	    	ELSE
		    	Response.Write nCount & " Items in Cart"
		    END IF
	    	Response.Write "</b></a>"
    	END IF
    %>
    </td>
  </tr>
</table>
</div>

<%
END SUB
cartDisplay


%>