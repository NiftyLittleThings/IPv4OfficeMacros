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

Function BitShiftWithWrapAround(cell as Double,shift As Integer)

	' We take a cell value, then shift with wraparound at 32 bits.
	' +ve shift moves right, -ve moves left.

	' After a huge effort fighting different type conversion (failures), gave up and implemented right shifts
	' as left shift by -32. Could get the upper 24 to rotate and not the lower 8 or the other way around.
	' Clng vs Double vs Int vs Fix vs / vs \ argh.

	Dim result as Double

   	If shift > 0 then ' we're right shifting
		
		result = BitShiftWithWrapAround(cell,shift-32)

	Else ' we shift left here
 
		Dim lresult as Variant
		Dim temp as Variant
		result = cell

		msb = 2^31
  
    		For x = 1 to abs(shift)

			If (result >= msb) Then
				' The MSB is set (1) - wrap to LSB (+1)
				result = ((result - msb) * 2) + 1
			Else
					' MSB is UNSET (0), roll left
				result = result * 2
			End If
		Next
		    
    	End If
	
	BitShiftWithWrapAround = result
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
