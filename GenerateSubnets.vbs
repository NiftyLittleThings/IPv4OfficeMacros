REM I GIVE UP TRYING TO GET THE INDENTATION TO WORK PROPERLY. FIXING IT SO IT LOOKS PERFECT SIMPLY REVERTS TO STRANGE JUNK AFTER THE SAVE.
REM AND I WON'T MENTION THE PATHETIC INDENTATION TAKING OVER EVERY TIME I HIT RETURN. SOME LINES ARE INDENTED 3/4 OF THE WAY ACROSS THE PAGE, WTF?
REM LIKE THE ONE BELOW - SHEESH

	rem
rem Helper functions to get values from other cells as String or Long
rem
function cellValAsString(sheet as String, col as Integer, row as Integer) as String
    cellValAsString = ThisComponent.getSheets().getByName(sheet).getCellByPosition(col,row).getString()
end function

function cellValAsVal(sheet as String, col as Integer, row as Integer) as Long
    cellValAsVal = ThisComponent.getSheets().getByName(sheet).getCellByPosition(col,row).getValue()
end function

sub GenerateSubnets()
    rem ----------------------------------------------------------------------
    rem Macro to generate all subnets of a given IP address, IF the CIDR block is known and some calcs have been done.
  
    rem Our setup is (it just is, don't ask ;):
    rem - the number of bits in the mask in J5
    rem - the IP in A3 and
    rem - Class A,B,C values corresponding to that IP (eg: 10.*,10.0.*,10.0.0.*) in cells H7, 8 and 9 - exercise for reader ;)
    rem - individual subnet values (from the subnet containing the IP address itself) decomposed into cells Q2:T2
    rem - The sheet name is Worksheet
    rem - the output is put into F11
	
	rem Example: 153.248.0.0/14 => generateSubnets() => '( 153.248.* OR 153.249.* OR 153.250.* OR 153.251.* )'
	rem This example will be used throughout
	
	rem Get the bit count. Find it's nearest 8 bit boundary. Sub that value.
    rem There will then be 2^remainder subnets to enumerate.

	Dim bitsInMask rem Example: /14
	Dim cellVal    rem Example: 153.248.0.0
	Dim debug
	debug = 0     rem It took quite some debugging, learning on the fly.
  
    Dim sheetName
    sheetName = "Worksheet"

	rem Eg 14
	bitsInMask = cellValueAsVal(sheetName,9,4) rem J5 (9,4)
	
	rem Eg 6
	remainder = bitsInMask mod 8

	rem This is the octet that we use later on - the one that doesn't change (and the ones to the left of it
	rem if there are any).
	
    rem Eg 1 We have a /14 so the first octet will remain static (offset 0) and the second one will be rotated (offset 1)
	octet = int(bitsInMask/8)
	
	if (remainder = 0) then
		rem We are either a /32 or expanding a /8, /16 or /24 so we can simply grab the
		rem existing strings from the Class A, B, C section, or the ip itself for the /32
		
		if (bitsInMask = 32) then rem Grab the cell itself
			qstr = cellValueAsString(sheetName,0,2) rem 0,2 => Cell A3
		else
			rem Grab one of the offsets in the Class A, B, C section - exercise for reader to generate these
			qstr = cellValueAsString(sheetName,5,9-octet)
		end if

	else
		rem If we are here then we need to expand a partial subnet into all it's constituent
		rem Elastic Search compatible string expressions - urgh. Worst cases are /9, /17, /25
		rem which all require 128 expansions (eg: 10/9 => 10.0.* -> 10.127.*)
		
		rem How many of these individual subnet strings will there be?
		
		rem Eg 2^(8-6) => 2^2 => 4 subnets. Correct for /14
		numSubs = 2^(8-remainder)
		if debug then msgbox "numSubs:" & numSubs
		
		rem We have the basis for the IP address at locations Q2:T2 so can determine which through
		rem the (offset + 16)
		rem Eg 17 
		offset = 16 + octet
		if debug then msgbox "Offset is: " & offset
	
		rem Set the start subnet value - this is the octet that is oscillated thru the range
		rem Eg R2, position 17,1 => 248
		startSub = cellValueAsVal(sheetName,offset,1)
		if debug then msgbox "As value: " & startSub
	
		rem Get each of the other subnet octets - we may not use them all but faster to set than to check which ones we need
		oct1 = cellValueAsString(sheetName,16,1)
		oct2 = cellValueAsString(sheetName,17,1)
		oct3 = cellValueAsString(sheetName,18,1)
		oct4 = cellValueAsString(sheetName,19,1)		
		if debug then msgbox oct1 & "." & oct2 & "." & oct3 & "." & oct4
		
		rem Set the end portion of the string which will be appended to each address - this is easy
		tail = ".*"
		
		rem Now set the starting portion of the string - need to determine which octet is cycling
		rem and only set it up to the preceeding octet including the trailing dot.  Then can cycle the next
        rem octet as required and keep gluing the trailer on if there is one to produce a full search term
		if (octet = 1) then
		
			rem Eg "153."
			head = oct1 & "."
	
		elseif (octet = 2) then
			head = oct1 & "." & oct2 & "."
	
		elseif (octet = 3) then
			head = oct1 & "." & oct2 & "." & oct3 & "."
			tail =""
            rem otherwise it adds a .* after the fully expanded /25-/31 addresses (ie:a bunch of /32's), which won't work as expected
			rem becasue the trailing dot doesn't match anything (the * is zero matched ok but the trailing dot is the problem)
			rem Note that the /32 case has been dealt with already above and the sub exited. The /32 case can never reach here.
	
		elseif (octet = 0) then
			head = ""
            rem special case for the multicast range 224.0.0.0/3
	
		else
			rem We should never see this - something is very wrong if you do
			msgbox "You should not be seeing this - something is very wrong if you do! Octet (0 based) not in range 0-3, value is: " & octet
			
			rem Now bail out stage left
			stop
	
		end if
		
		if debug then msgbox "Head is: " & head & chr(13) & chr(13) & "Tail is: " & tail

		rem Init the query string - we're building for ElasticSearch so we want to gather the terms in ()'s
        rem as they are ultimately preceeded by more prefixes, ala:
        rem '...{ query_string: { query = "source_ip: (our generated subnets..)" } } ...'
		
        qstr = "( "
		
        rem Build up the query string one subnet at a time with OR between each, ala for 153.248.0.0/14:
		rem ( 153.248.* OR 153.249.* OR 153.250.* OR 153.251.* )
		rem At this point, qstr="( ", head="153.", tail=".*", startSub=248
    
        for x = 0 to (numSubs-2)
			rem We do 2 cycles less because we started at 0 offset from starting sub and special case for last one so the extra OR is not appended
						
			qstr = qstr & head & (x+startSub) & tail & " OR "
		next
		
		rem Ok now glue the end onto what we have built so far, which is all the subnets except for the very last one
		rem Eg qstr="( 153.248.* OR 153.249.* OR 153.250.* OR " & head "153." & (248+4-1)=251 & tail ".*" & ")"
		
        qstr = qstr & head & (startSub + numSubs-1) & tail & " ) "

        if debug then msgbox qstr

    end if

    rem Set the cell value into F11
    rem F11 => 5,10
    ThisComponent.getSheets().getByName(sheetName).getCellByPosition(5,10).setString(qstr)
  
    rem Why not make this a function and return the value? Because we can override content of F11 and it's convenient
    rem to have this function replace it as required with no user interaction. One click.  We could make a wrapper for it
    rem but it's only used a few times and the dependency in those cases is to have the result in F11 (for _our_ purposes).

    rem Ok, so we made it easy for you and built a custom sheet just to make it easy. Unfortunately that code differs from
    rem this version with all these extra comments. But see SubnetExpansionGenerator.ods anyway.
end sub
