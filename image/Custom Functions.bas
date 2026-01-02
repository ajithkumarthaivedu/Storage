Attribute VB_Name = "Module11"
Function SmartVLookupLocked(lookupValue As Variant, tableRange As Range, Optional headerCell As Range, Optional matchType As Variant) As Variant
    Dim colIndex As Long
    Dim lockedTable As Range
    Dim lockedHeader As Range

    ' Force absolute references by re-resolving the address
    Set lockedTable = Range(tableRange.Address(External:=True))
    If Not headerCell Is Nothing Then Set lockedHeader = Range(headerCell.Address(External:=True))

    ' Determine column index
    If Not headerCell Is Nothing Then
        On Error GoTo HeaderNotFound
        colIndex = Application.WorksheetFunction.Match(lockedHeader.Value, lockedTable.Rows(1), 0)
    Else
        colIndex = lockedTable.Columns.Count
    End If

    ' Default matchType
    If IsMissing(matchType) Then matchType = False

    ' Perform VLOOKUP
    SmartVLookupLocked = Application.WorksheetFunction.VLookup(lookupValue, lockedTable, colIndex, matchType)
    Exit Function

HeaderNotFound:
    SmartVLookupLocked = CVErr(xlErrNA)
End Function



Public Function Revisit_Conversion( _
    leaddatev As Variant, leadnumv As Variant, _
    leaddate As Range, leadnum As Range, _
    apptdate As Range, apptnum As Range, _
    Optional ApptStatus As Range = Nothing _
) As Variant

    Dim MinApptDate As Variant, LatestLeadDate As Variant

    If Not ApptStatus Is Nothing Then
        ' With status filter
        MinApptDate = Application.MinIfs(apptdate, apptnum, leadnumv, apptdate, ">=" & CLng(leaddatev), ApptStatus, "Paid")
        If IsError(MinApptDate) Or MinApptDate = 0 Then Revisit_Conversion = "-": Exit Function

        LatestLeadDate = Application.MaxIfs(leaddate, leadnum, leadnumv, leaddate, "<=" & CLng(MinApptDate))
    Else
        ' Without status filter
        MinApptDate = Application.MinIfs(apptdate, apptnum, leadnumv, apptdate, ">=" & CLng(leaddatev))
        If IsError(MinApptDate) Or MinApptDate = 0 Then Revisit_Conversion = "-": Exit Function

        LatestLeadDate = Application.MaxIfs(leaddate, leadnum, leadnumv, leaddate, "<=" & CLng(MinApptDate))
    End If

    ' Compare lead date with latest lead date
    If leaddatev = LatestLeadDate Then
        Revisit_Conversion = MinApptDate
    Else
        Revisit_Conversion = "-"
    End If

End Function


Function bedbooking(Date_i As Variant, UHID As Variant, tableRange As Range) As Variant
    Dim i As Integer
    Dim Date_i1 As Variant
    Dim a As Variant

    On Error Resume Next ' Enable error handling
    For i = 1 To 9
        ' Calculate the first day of the previous month
        Date_i1 = Application.WorksheetFunction.EoMonth(Date_i, -i) + 1
        
        ' Perform VLookup to find the value
        a = Application.WorksheetFunction.VLookup(Date_i1 & UHID, tableRange, 2, False)
        
        ' Check if an error occurred
        If Err.Number <> 0 Then
            Err.Clear ' Clear the error and continue
        Else
            bedbooking = a ' Assign the found value to the function
            Exit Function ' Exit the function
        End If
    Next i
    On Error GoTo 0 ' Disable error handling
End Function


Function FY(d As Date) As Variant
    'Dim fiscalYear As Integer
    If Month(d) >= 4 Then
        fiscalYear = "FY " & Right(Year(d), 2) + 1
    Else
        fiscalYear = "FY " & Right(Year(d), 2)
    End If
    FY = fiscalYear
End Function

Function Manpower(Calls As Variant, Duration As Variant, Occ As Variant, AHT As Variant, Optional shr As Variant) As Variant
If shr = 0 Then
Manpower = Application.WorksheetFunction.RoundUp((Calls * AHT) / (Occ * Duration * 86400), 0)
Else
Manpower = Application.WorksheetFunction.RoundUp(((Calls * AHT) / (Occ * Duration * 86400)) / (1 - shr), 0)
End If
End Function


Function ExtDate(a As Variant) As Variant
    ExtDate = CDate(Left(a, 10))
End Function


Function InvoiceDate(CRM_No As Variant, CRM_Date As Variant, Appt_No As Range, Appt_Date As Range) As Variant

r1 = Application.WorksheetFunction.Match(CRM_No, Appt_No, 0)
r2 = Application.WorksheetFunction.Match(CRM_No, Appt_No, 1)

    For i = r1 To r2
        If Appt_Date.Cells(i, 1) >= CRM_Date Then
            InvoiceDate = Appt_Date.Cells(i, 1)
            Exit Function
        End If
    Next i
'ApptDate = r1 & "," & r2
End Function

Function apptdate(CRM_No As Variant, CRM_Date As Variant, Appt_No As Range, Appt_Date As Range) As Variant

r1 = Application.WorksheetFunction.Match(CRM_No, Appt_No, 0)
r2 = Application.WorksheetFunction.Match(CRM_No, Appt_No, 1)

    For i = r1 To r2
        If Appt_Date.Cells(i, 1) >= CRM_Date Then
            apptdate = Appt_Date.Cells(i, 1)
            Exit Function
        End If
    Next i
'ApptDate = r1 & "," & r2
End Function


Function a(rng As Range, val As Variant, Optional er As Variant) As Variant
If Application.WorksheetFunction.Match(val, rng, 0) >= val.Row - rng.Rows(1).Row + 1 Then a = 1
End Function

Function aa(rng As Range, val As Range) As Variant

Dim rs1 As Integer, rs2 As Integer, rs3 As Integer, rs4 As Integer
r = val.Row

'1
On Error Resume Next
If Application.WorksheetFunction.Match(rng.Cells(r, 1), rng.Columns(1), 0) >= val.Row - rng.Rows(1).Row + 1 Then
rs1 = 1
Else
rs1 = 0
End If

'2
On Error Resume Next
If Application.WorksheetFunction.Match(rng.Cells(r, 1), rng.Columns(2), 0) >= val.Row - rng.Rows(1).Row + 1 Then
rs2 = 1
Else
rs2 = 0
End If

'3
On Error Resume Next
If Application.WorksheetFunction.Match(rng.Cells(r, 2), rng.Columns(2), 0) >= val.Row - rng.Rows(1).Row + 1 Then
rs3 = 1
Else
rs3 = 0
End If

'4
On Error Resume Next
If Application.WorksheetFunction.Match(rng.Cells(r, 2), rng.Columns(1), 0) >= val.Row - rng.Rows(1).Row + 1 Then
rs4 = 1
Else
rs4 = 0
End If

     aa = Application.WorksheetFunction.Min(rs1, rs2, rs3, rs4)
End Function


Function Repeats(mob_Rng As Range, num As Variant, date_rng As Range, dt As Variant, n As Variant) As Variant
Dim a As Integer

For i = 0 To n - 1
Count = Application.WorksheetFunction.CountIfs(mob_Rng, num, date_rng, dt - i) + a
a = Count
Next i

If a > 1 Then a = 1 Else a = 0
Repeats = a

End Function


Function Split1(Text As String, de As String) As Variant
Dim arr() As String, textc As String
textc = Application.WorksheetFunction.Trim(Text)
arr = Split(textc, de)
For i = LBound(arr) To UBound(arr)
arr(i) = Trim(arr(i))
Next i

Split1 = arr
End Function


Function VMatch(lookupValue As Variant, Header As Variant, tableRange As Range, Optional er As Variant) As Variant
r = tableRange.Row

On Error Resume Next
c = Application.WorksheetFunction.Match(Header, tableRange.Rows(1), 0)
    If Err.Number <> 0 Then
    'VMatch = VBA.CVErr(xlErrNA)
    VMatch = er
    Err.Clear
    Exit Function
    End If

On Error Resume Next
VMatch = Application.WorksheetFunction.VLookup(lookupValue, tableRange, c, 0)
    If Err.Number <> 0 Then
    'VMatch = VBA.CVErr(xlErrNA)
    VMatch = er
    Err.Clear
    End If

End Function




Function MobileNumber(mobile_no As Variant, Optional iferror As Variant) As Variant
On Error Resume Next
b = Application.WorksheetFunction.Find(",", mobile_no, 1)
c = Application.WorksheetFunction.Find("/", mobile_no, 1)
d = Application.WorksheetFunction.Find("\", mobile_no, 1)
e = Application.WorksheetFunction.Find("-", mobile_no, 1)
If e > 10 Then e = e Else e = 0

b = Application.WorksheetFunction.Max(b, c, d, e)

If b <> 0 Then
    mobile_no = Left(mobile_no, b - 1)
End If

    For i = 1 To Len(mobile_no)
        If Mid(mobile_no, i, 1) Like "[0-9]" Then
            Nu = Nu & Mid(mobile_no, i, 1)
        End If
    Next i
    
    Nu = Nu * 1
    
    If Len(Nu) > 10 And Left(Nu, 2) = 91 Then
    Nu = Right(Nu, 10) * 1
    End If
    
    If Nu < 9999999999# And Nu > 6000000000# Then
    Nu = Nu
    ElseIf iferror = 0 Then
    Nu = CVErr(xlErrNA)
    Else
    Nu = iferror
    End If
    
    MobileNumber = Nu
    

End Function


Function V2Range(lookupValue As Variant, LookupValue2 As Variant, tableRange As Range, TableRange2 As Range, Optional iferror As Variant) As Variant

c = tableRange.Columns.Count
c2 = TableRange2.Columns.Count

    On Error Resume Next
    
                            V2Range = Application.WorksheetFunction.VLookup(lookupValue, tableRange, c, 0)
    If Err.Number <> 0 Then V2Range = Application.WorksheetFunction.VLookup(lookupValue, TableRange2, c2, 0)
    If Err.Number <> 0 Then V2Range = Application.WorksheetFunction.VLookup(LookupValue2, tableRange, c, 0)
    If Err.Number <> 0 Then V2Range = Application.WorksheetFunction.VLookup(LookupValue2, TableRange2, c2, 0)
    
    If V2Range = 0 And Err.Number <> 0 Then
    If iferror = 0 Then V2Range = CVErr(xlErrNA) Else V2Range = iferror
    End If
            
End Function

Function FindInstance(cellvalue As Variant, rng As Range, Optional iferror As Variant) As Variant

Dim i As Integer
Dim pos As Integer
Dim arr As Range
Dim a As String

Set arr = rng
cellvalue1 = Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(cellvalue))

arr1 = Application.Transpose(rng)
    
    For i = LBound(arr1) To UBound(arr1)
    On Error Resume Next
    
    If arr1(i) = "" Then a = "abcd123" Else a = arr1(i)
    
        pos = InStr(1, cellvalue1, a, vbTextCompare)
        If pos > 0 Then
            FindInstance = arr1(i)
            Exit Function
        End If
    Next i
    
    If FindInstance = 0 Then
    If iferror = 0 Then FindInstance = CVErr(xlErrNA) Else FindInstance = iferror
    End If
    
End Function
Function Find_And_Match(cellvalue As Variant, rng As Range, Optional iferror As Variant) As Variant
Dim a As Variant
a = FindInstance(cellvalue, rng.Columns(1))

If Err.Number <> 0 Then
If iferror <> 0 Then Find_And_Match = iferror
Err.Clear
Exit Function
End If

Find_And_Match = Application.WorksheetFunction.VLookup(a, rng, rng.Columns.Count, 0)

End Function


Function GetColumnRange(lookupValue As Variant, Header As Range) As String

r = Header.Row
c = Application.WorksheetFunction.Match(lookupValue, Header, 0)
a = Split(Cells(r, c).Address, "$")(1)
b = Range(a & ":" & a).Address

GetColumnRange = b

'GetColumnRange = a & ":" & a
End Function

Function Correct(a)
   For i = 1 To Len(a)
        If Mid(a, i, 1) Like "[0-9]" Then
            Nu = Nu & Mid(a, i, 1)
        End If
    Next i
 Correct = Nu * 1
End Function

Function Form()
UserForm1.Show
End Function
Function GetSheetNames() As Variant
    Dim i As Long
    Dim wsCount As Long
    wsCount = ThisWorkbook.Sheets.Count
    Dim result() As Variant
    ReDim result(1 To wsCount, 1 To 1)

    For i = 1 To wsCount
        result(i, 1) = ThisWorkbook.Sheets(i).Name
    Next i

    GetSheetNames = result
End Function



















