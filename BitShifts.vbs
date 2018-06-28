rem Function to call a built-in LibreOffice function
Function calc_Func(sFunc$,args())
	dim oFA as Object
	oFA = createUNOService("com.sun.star.sheet.FunctionAccess")
	calc_Func = oFA.callFunction(sFunc,args())
end function


Function BitNotStr(cellval) as string
rem Function to walk a string of 1s and 0s and invert them all
rem
rem If A10 contains '11110000' then BitNotStr(A10) = '00001111'
rem
rem Implemented using a crude loop calling MID on the string repetitively.
rem
rem Seems stupid but this version handles much longer strings than the internal version - hence it's existence :)

	dim loops as integer
	dim temp as string
	
	loops = len(cellval)
	for i=1 to loops 
		thebit = calc_Func("MID",array(cellval,i,1))
		if (thebit = "1") then
			temp=temp & "0"
		else
			temp=temp & "1"
		end if
	next

	bitnot = temp

End Function

Function BitNotDirect(cellval) as long
rem Takes a numeric value and directly implements the inversion functions - yuck, y intentional ;)

	dim tempstr as string
	dim radix, bitLength as integer
	
	radix = 2
	bitLength = 8
	
	tempstr = calc_Func("BASE",array(cellval,radix,bitLength))
	
	tempstr = BitNot(tempstr)
	
	dim ret as long
	
	ret = calc_Func("BIN2DEC",array(tempstr))
	
	BitNotDirect = ret
End Function

Function BitShiftWithOverflow(cell,shift,length As Integer)
rem Unimplemented - merely a pass through to the default BitLShift built-in function

    Dim result As Variant
	
    If NOT IsMissing(cell) Then
    	result = calc_Func("BITLSHIFT",array(cell,shift))
    End If

    BitShiftWithOverflow = result
End Function
