rem
rem A String Join function. Takes a range and joins it together into one long string separated by optional
rem joining, before and after strings.

rem i.e.: StrJoin(range,delimiter,before,after)

rem range = 'A2:A10' => IP address 1, IP address 2, etc
rem StrJoin(range, ' OR ', '( ',  ' )' ) => '( IP address 1 OR IP address 2 OR ... )'


rem Why do we want this?

rem Graylog queries via parameters placed in the URL, which is limited in length by URL standards. 

rem As someone else found out:
rem https://github.com/Graylog2/graylog2-server/issues/3317

rem But we can POST to ElasticSearch directly with a body of (effectively for our purposes) unlimited length.
rem Just not via Graylog's UI.

rem However, we need to construct our query from data stored in the spreadsheet.

rem This function makes it trivial to join together a list of IP's copy and pasted from IoC sources.
rem Or tcp sequence numbers... or bit shifted things

rem Another example: AA6:AA37 contain bit shifted IP addresses
rem = STRJOIN(Sheet7.AA6:AA37," OR ")
rem = STRJOIN(Sheet7.AA6:AA37,"OR"," "," ") achieves same result
rem 21.183.0.1 OR 21.183.0.2 OR 21.183.0.4 OR 21.183.0.8 OR 21.183.0.16 OR 21.183.0.32 OR 21.183.0.64 OR
rem 21.183.0.128 OR 21.183.1.0 OR 21.183.2.0 OR 21.183.4.0 OR 21.183.8.0 OR 21.183.16.0 OR 21.183.32.0 OR
rem 21.183.64.0 OR 21.183.128.0 OR 21.182.0.0 OR 21.181.0.0 OR 21.179.0.0 OR 21.191.0.0 OR 21.167.0.0 OR
rem 21.151.0.0 OR 21.247.0.0 OR 21.55.0.0 OR 20.183.0.0 OR 23.183.0.0 OR 17.183.0.0 OR 29.183.0.0 OR 5.183.0.0
rem OR 53.183.0.0 OR 85.183.0.0 OR 149.183.0.0

rem Imagine typing that! One click, I say...

Function STRJOIN(range, Optional delimiter As String, Optional before As String, Optional after As String)
    Dim row, col As Integer
    Dim result, cell As String

    result = ""

    If IsMissing(delimiter) Then
        delimiter = ","
    End If
    If IsMissing(before) Then
        before = ""
    End If
    If IsMissing(after) Then
        after = ""
    End If

    If NOT IsMissing(range) Then
    
        rem Single cell case
        If NOT IsArray(range) Then
            result = before & range & after
        Else
            For row = LBound(range, 1) To UBound(range, 1)
                For col = LBound(range, 2) To UBound(range, 2)
                    cell = range(row, col)
                    If cell <> 0 AND Len(Trim(cell)) <> 0 Then
                        If result <> "" Then
                            result = result & delimiter
                        End If
                        result = result & before & range(row, col) & after
                    End If
                Next
            Next
        End If
    End If

    STRJOIN = result
End Function
