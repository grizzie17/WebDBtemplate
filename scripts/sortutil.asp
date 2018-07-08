<%
'+------------------------------------------------------------------+
'|
'|			  Copyright 1986 .. 2004 Grizzly WebMaster
'|				a division of Bear Consulting Group
'|						  All Rights Reserved
'|
'|	This software-file/document, in whole or in part, including
'|	the structures and the procedures described herein, may not
'|	be provided or otherwise made available without prior written
'|	authorization.  In case of authorized or unauthorized
'|	publication or duplication, copyright is claimed.
'|
'+------------------------------------------------------------------+




' qsort

Sub QSort( ByRef A(), ByVal Lb, ByVal Ub )
    Dim lbStack(32)
    Dim ubStack(32)
    Dim sp              ' stack pointer
    Dim lbx             ' current lower-bound
    Dim ubx             ' current upper-bound
    Dim m
    Dim p               ' index to pivot
    Dim i
    Dim j
    Dim t            ' temp used for exchanges

    lbStack(0) = Lb
    ubStack(0) = Ub
    sp = 0
    Do While sp >= 0
        lbx = lbStack(sp)
        ubx = ubStack(sp)

        Do While (lbx < ubx)

            ' select pivot and exchange with 1st element
            p = lbx + (ubx - lbx) \ 2

            ' exchange lbx, p
            t = A(lbx)
            A(lbx) = A(p)
            A(p) = t

            ' partition into two segments
            i = lbx + 1
            j = ubx
            Do
                Do While i < j
                    If A(lbx) <= A(i) Then Exit Do
                    i = i + 1
                Loop

                Do While j >= i
                    If A(j) <= A(lbx) Then Exit Do
                    j = j - 1
                Loop

                If i >= j Then Exit Do

                ' exchange i, j
                t = A(i)
                A(i) = A(j)
                A(j) = t

                j = j - 1
                i = i + 1
            Loop

            ' pivot belongs in A[j]
            ' exchange lbx, j
            t = A(lbx)
            A(lbx) = A(j)
            A(j) = t

            m = j

            ' keep processing smallest segment, and stack largest
            If m - lbx <= ubx - m Then
                If m + 1 < ubx Then
                    lbStack(sp) = m + 1
                    ubStack(sp) = ubx
                    sp = sp + 1
                End If
                ubx = m - 1
            Else
                If m - 1 > lbx Then
                    lbStack(sp) = lbx
                    ubStack(sp) = m - 1
                    sp = sp + 1
                End If
                lbx = m + 1
            End If
        Loop
        sp = sp - 1
    Loop
End Sub


Sub QuickSort( vec, loBound, hiBound )
	Dim pivot,loSwap,hiSwap,temp

	'== This procedure is adapted from the algorithm given in:
	'==    Data Abstractions & Structures using C++ by
	'==    Mark Headington and David Riley, pg. 586
	'== Quicksort is the fastest array sorting routine for
	'== unordered arrays.  Its big O is  n log n


	'== Two items to sort
	if hiBound - loBound = 1 then
		if vec(loBound) > vec(hiBound) then
			temp=vec(loBound)
			vec(loBound) = vec(hiBound)
			vec(hiBound) = temp
		End If
	End If

	'== Three or more items to sort
	pivot = vec(int((loBound + hiBound) / 2))
	vec(int((loBound + hiBound) / 2)) = vec(loBound)
	vec(loBound) = pivot
	loSwap = loBound + 1
	hiSwap = hiBound
  
	do
		'== Find the right loSwap
		while loSwap < hiSwap and vec(loSwap) <= pivot
			loSwap = loSwap + 1
		wend
		'== Find the right hiSwap
		while vec(hiSwap) > pivot
			hiSwap = hiSwap - 1
		wend
		'== Swap values if loSwap is less then hiSwap
		if loSwap < hiSwap then
			temp = vec(loSwap)
			vec(loSwap) = vec(hiSwap)
			vec(hiSwap) = temp
		End If
	loop while loSwap < hiSwap
  
	vec(loBound) = vec(hiSwap)
	vec(hiSwap) = pivot
  
	'== Recursively call function .. the beauty of Quicksort
	'== 2 or more items in first section
	if loBound < (hiSwap - 1) then Call QuickSort(vec,loBound,hiSwap-1)
	'== 2 or more items in second section
	if hiSwap + 1 < hibound then Call QuickSort(vec,hiSwap+1,hiBound)

End Sub  'QuickSort





SUB sort( A(), ByVal Lb, ByVal Ub )
    DIM n
    DIM h
    DIM i
    DIM j
    DIM t

    ' sort array[lb..ub]

    ' compute largest increment
    n = Ub - Lb + 1
    h = 1
    IF (n < 14) THEN
        h = 1
    ELSE
        DO WHILE h < n
            h = 3 * h + 1
        LOOP
        h = h \ 3
        h = h \ 3
    END IF

    DO WHILE h > 0
        ' sort by insertion in increments of h
        FOR i = Lb + h TO Ub
            t = A(i)
            FOR j = i - h TO Lb STEP -h
                IF A(j) <= t THEN EXIT FOR
                A(j + h) = A(j)
            NEXT 'j
            A(j + h) = t
        NEXT 'i
        h = h \ 3
    LOOP
END SUB


SUB dualSort( A(), B(), ByVal Lb, ByVal Ub )
    DIM n
    DIM h
    DIM i
    DIM j
    DIM t
	DIM	t2

    ' sort array[lb..ub]

    ' compute largest increment
    n = Ub - Lb + 1
    h = 1
    IF (n < 14) THEN
        h = 1
    ELSE
        DO WHILE h < n
            h = 3 * h + 1
        LOOP
        h = h \ 3
        h = h \ 3
    END IF

    DO WHILE h > 0
        ' sort by insertion in increments of h
        FOR i = Lb + h TO Ub
            t = A(i)
			t2 = B(i)
            FOR j = i - h TO Lb STEP -h
                IF A(j) <= t THEN EXIT FOR
                A(j + h) = A(j)
				B(j + h) = B(j)
            NEXT 'j
            A(j + h) = t
			B(j + h) = t2
        NEXT 'i
        h = h \ 3
    LOOP
END SUB


SUB dualSortDescend( A(), B(), ByVal Lb, ByVal Ub )
    DIM n
    DIM h
    DIM i
    DIM j
    DIM t
	DIM	t2

    ' sort array[lb..ub]

    ' compute largest increment
    n = Ub - Lb + 1
    h = 1
    IF (n < 14) THEN
        h = 1
    ELSE
        DO WHILE h < n
            h = 3 * h + 1
        LOOP
        h = h \ 3
        h = h \ 3
    END IF

    DO WHILE h > 0
        ' sort by insertion in increments of h
        FOR i = Lb + h TO Ub
            t = A(i)
			t2 = B(i)
            FOR j = i - h TO Lb STEP -h
                IF t <= A(j) THEN EXIT FOR
                A(j + h) = A(j)
				B(j + h) = B(j)
            NEXT 'j
            A(j + h) = t
			B(j + h) = t2
        NEXT 'i
        h = h \ 3
    LOOP
END SUB



SUB sortDescend( A(), ByVal Lb, ByVal Ub )
    DIM n
    DIM h
    DIM i
    DIM j
    DIM t
	DIM	t2

    ' sort array[lb..ub]

    ' compute largest increment
    n = Ub - Lb + 1
    h = 1
    IF (n < 14) THEN
        h = 1
    ELSE
        DO WHILE h < n
            h = 3 * h + 1
        LOOP
        h = h \ 3
        h = h \ 3
    END IF

    DO WHILE h > 0
        ' sort by insertion in increments of h
        FOR i = Lb + h TO Ub
            t = A(i)
            FOR j = i - h TO Lb STEP -h
                IF t <= A(j) THEN EXIT FOR
                A(j + h) = A(j)
            NEXT 'j
            A(j + h) = t
        NEXT 'i
        h = h \ 3
    LOOP
END SUB


SUB shuffle( A(), ByVal Lb, ByVal Ub )

	RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
	
	DIM	i
	DIM	k
	DIM	j
	DIM t
	FOR i = Lb TO Ub
		k = Ub - i
		IF 1 < k THEN
			j = INT( k * RND + 0.5 ) + i
			IF ub < j THEN j = ub
			t = A(i)
			A(i) = A(j)
			A(j) = t
		END IF
	NEXT 'i
	
END SUB





%>
