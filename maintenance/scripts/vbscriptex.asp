<%
' these are some simple functions that should have been included in the language



FUNCTION min( x, y )

	IF x < y THEN
		min = x
	ELSE
		min = y
	END IF

END FUNCTION


FUNCTION max( x, y )

	IF x < y THEN
		max = y
	ELSE
		max = x
	END IF
	
END FUNCTION


%>