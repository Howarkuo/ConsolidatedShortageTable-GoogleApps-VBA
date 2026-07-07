Sub AddStatusColumnAfterP()
    ' Insert a new column at Q, shifting existing columns to the right
    Columns("Q:Q").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Set the header name in the first row of the new column
    Range("Q4").Value = "STATUS"
End Sub

Sub ExpandQRowsToDemandSupplyBalance()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    
    Set ws = ActiveSheet
    
    ' Based on this ALL-WIP-COMP-Less table, data starts at row 5. Adjust if needed.
    startRow = 5
    
    ' Find the last used row based on Column P
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    
    


    ' Turn off screen updating to make the macro run much faster
    Application.ScreenUpdating = False
    
    ' Loop backwards from the last row up to the start row
    For i = lastRow To startRow Step -1
        ' Insert 2 rows directly below the current row
        ws.Rows(i + 1 & ":" & i + 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        ' OPTIONAL: If you want to copy the item info (Columns A to P) down to the new rows,
        ' remove the single quote (') from the beginning of the next line:
        ' ws.Range(ws.Cells(i, 1), ws.Cells(i, 16)).Copy Destination:=ws.Range(ws.Cells(i + 1, 1), ws.Cells(i + 2, 16))
        
        ' Set the values in Column Q (Column 17)
        ws.Cells(i, 17).Value = "DEMAND"
        ws.Cells(i + 1, 17).Value = "SUPPLY"
        ws.Cells(i + 2, 17).Value = "BALANCE"
    Next i
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    MsgBox "Rows expanded successfully!", vbInformation
End Sub


Sub SupplyWriteIn()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 5 Then Exit Sub ' Nothing to process
    
    ' --- STEP 1: PARSE HEADERS ONCE ---
    Dim colCount As Long
    colCount = 43 - 18 + 1 ' 26 columns
    
    Dim hStart() As Date
    Dim hEnd() As Date
    ReDim hStart(1 To colCount)
    ReDim hEnd(1 To colCount)
    
    ' NEW: Collection acts as a Mac-compatible Dictionary for O(1) exact date lookups
    Dim exactDates As New Collection
    
    Dim c As Long
    Dim hVal As String
    Dim rawVal As Variant
    Dim p() As String
    Dim parsedDate As Date
    
    For c = 1 To colCount
        ' Using .Value fixes the "#######" visual width issue
        rawVal = ws.Cells(4, c + 17).Value
        hVal = Trim(CStr(rawVal))
        
        If InStr(hVal, "~") > 0 Then
            p = Split(hVal, "~")
            hStart(c) = ParseMD(p(0))
            hEnd(c) = ParseMD(p(1))
        ElseIf Left(hVal, 4) = "Over" Then
            hStart(c) = ParseMD(Replace(hVal, "Over", "")) + 1
            hEnd(c) = DateSerial(2099, 12, 31)
        Else
            ' Determine exact date
            If IsDate(rawVal) Then
                parsedDate = CDate(rawVal)
            Else
                parsedDate = ParseYMD(hVal)
            End If
            
            hStart(c) = parsedDate
            hEnd(c) = parsedDate
            
            ' Add to our O(1) lookup collection (using the date's underlying number as the Key)
            If parsedDate > 0 Then
                On Error Resume Next ' Prevents error if multiple identical headers exist
                exactDates.Add Item:=c, Key:=CStr(CDbl(parsedDate))
                On Error GoTo 0
            End If
        End If
    Next c
    
    ' --- STEP 2: LOAD DATA INTO MEMORY FOR SPEED ---
    Dim inF As Variant
    inF = ws.Range("F5:F" & lastRow).Value
    
    Dim outVals As Variant
    ' FIXED: Exact array bound mapping (No +1 needed)
    outVals = ws.Range("R5").Resize(UBound(inF, 1), colCount).Value
    
    ' --- STEP 3: PROCESS DATA IN MEMORY ---
    Dim r As Long, i As Long
    Dim strF As String
    Dim parts() As String, itemParts() As String
    Dim eDate As Date, qty As Double
    Dim matchCol As Long
    
    ' Start at index 1 (Row 5) and step by 3 (Rows 5, 8, 11...)
    For r = 1 To UBound(inF, 1) Step 3
        
        If Not IsError(inF(r, 1)) Then
            strF = Trim(CStr(inF(r, 1)))
            
            If Len(strF) > 0 Then
                parts = Split(strF, ";")
                
                For i = LBound(parts) To UBound(parts)
                    itemParts = Split(Trim(parts(i)), "*")
                    
                    If UBound(itemParts) = 1 Then
                        eDate = ParseYMD(itemParts(0))
                        qty = Val(itemParts(1))
                        
                        If qty <> 0 And eDate > 0 Then
                            
                            ' --- O(1) HASH MAP LOOKUP ---
                            matchCol = 0
                            On Error Resume Next
                            matchCol = exactDates(CStr(CDbl(eDate)))
                            On Error GoTo 0
                            
                            If matchCol > 0 Then
                                ' INSTANT MATCH: Date found in exact headers
                                If IsEmpty(outVals(r + 1, matchCol)) Or outVals(r + 1, matchCol) = "" Then
                                    outVals(r + 1, matchCol) = qty
                                Else
                                    outVals(r + 1, matchCol) = outVals(r + 1, matchCol) + qty
                                End If
                            Else
                                ' FALLBACK: Only loop for ranges (e.g., 07/09~07/15)
                                For c = 1 To colCount
                                    If eDate >= hStart(c) And eDate <= hEnd(c) Then
                                        If IsEmpty(outVals(r + 1, c)) Or outVals(r + 1, c) = "" Then
                                            outVals(r + 1, c) = qty
                                        Else
                                            outVals(r + 1, c) = outVals(r + 1, c) + qty
                                        End If
                                        Exit For
                                    End If
                                Next c
                            End If
                            
                        End If
                    End If
                Next i
            End If
        End If
    Next r
    
    ' --- STEP 4: WRITE BACK TO SHEET IN ONE GO ---
    ws.Range("R5").Resize(UBound(outVals, 1), colCount).Value = outVals

    MsgBox "Quantities assigned successfully!", vbInformation
End Sub



' --- HELPER FUNCTIONS ---

Function ParseMD(ByVal txt As String) As Date
    On Error Resume Next
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, vbCr, "")
    
    Dim p() As String
    p = Split(Trim(txt), "/")
    If UBound(p) = 1 Then
        ParseMD = DateSerial(2026, CLng(p(0)), CLng(p(1)))
    End If
    On Error GoTo 0
End Function

Function ParseYMD(ByVal txt As String) As Date
    On Error Resume Next
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, vbCr, "")
    
    Dim p() As String
    p = Split(Trim(txt), "/")
    If UBound(p) = 2 Then
        ParseYMD = DateSerial(CLng(p(0)), CLng(p(1)), CLng(p(2)))
    End If
    On Error GoTo 0
End Function

Sub FillBlanksWithZero()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    
    If lastRow < 5 Then Exit Sub

    ' Turn off error handling temporarily just in case there are NO blank cells
    ' (otherwise SpecialCells throws an error)
    On Error Resume Next
    
    ' Selects all blank cells between R5 and AQ[lastRow] and fills them with 0 instantly
    ws.Range("R5:AR" & lastRow).SpecialCells(xlCellTypeBlanks).Value = 0
    
    On Error GoTo 0
    
    MsgBox "All blank cells filled with 0!", vbInformation
End Sub

Sub CalculateBalances()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    ' Look at Column Q ("STATUS") to find the true last row
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 7 Then Exit Sub ' Needs to be at least row 7 to have a BALANCE row
    
    Dim colCount As Long
    colCount = 44 - 18 + 1 ' 27 columns (R to AR)
    
    ' --- STEP 1: LOAD INTO MEMORY ---
    Dim arr As Variant
    arr = ws.Range("R5").Resize(lastRow - 4, colCount).Value
    
    Dim r As Long, c As Long
    Dim demand As Double, supply As Double, prevBalance As Double
    
    ' --- STEP 2: CALCULATE IN MEMORY ---
    ' Loop starting at array index 3 (Row 7), stepping by 3 (Rows 7, 10, 13...)
    For r = 3 To UBound(arr, 1) Step 3
        
        ' 1. Calculate the very first column (Column R / Index 1)
        ' Since there is no previous column, Balance = Demand + Supply
        demand = Val(CStr(arr(r - 2, 1)))
        supply = Val(CStr(arr(r - 1, 1)))
        arr(r, 1) = demand + supply
        
        ' 2. Calculate the rolling balance for the rest of the columns (S to AR)
        For c = 2 To colCount
            prevBalance = Val(CStr(arr(r, c - 1))) ' Previous column's balance
            demand = Val(CStr(arr(r - 2, c)))      ' Current column's demand
            supply = Val(CStr(arr(r - 1, c)))      ' Current column's supply
            
            arr(r, c) = prevBalance + demand + supply
        Next c
    Next r
    
    ' --- STEP 3: WRITE BACK IN ONE GO ---
    ws.Range("R5").Resize(UBound(arr, 1), colCount).Value = arr
    
    MsgBox "Balances calculated perfectly!", vbInformation
End Sub

Sub HighlightNegativeBalances()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRange As Range
    
    Set ws = ActiveSheet
    
    ' Look at Column Q ("STATUS") to find the true last row
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 7 Then Exit Sub ' Needs to be at least row 7 to have a BALANCE row
    
    ' Define the entire block of data from R7 to the bottom right of AQ
    Set targetRange = ws.Range("R7:AR" & lastRow)
    
    ' --- THE O(1) SPEED TRICK ---
    ' Clear any old formatting rules so they don't stack up if you run this twice
    targetRange.FormatConditions.Delete
    
    ' Apply a single rule to the entire grid at once.
    ' Logic: "If the value is < 0 AND the Row Number divided by 3 has a remainder of 1 (Rows 7, 10, 13...)"
    With targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(R7<0, MOD(ROW(),3)=1)")
        .Font.Color = vbRed
        ' Optional: Make it bold so it stands out more!
        .Font.Bold = True
    End With
    
    MsgBox "Negative balances are now red!", vbInformation
End Sub

Sub FillDown_ItemNo()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    
    ' If the last row is less than 5, there's nothing to process
    If lastRow < 5 Then Exit Sub
    
    ' Set the range to Column A starting from row 5
    Set rng = ws.Range("A5:A" & lastRow)
    
    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    
    On Error Resume Next ' Prevent error if there are no blank cells
    With rng.SpecialCells(xlCellTypeBlanks)
        ' Insert formula to equal the cell exactly one row above
        .FormulaR1C1 = "=R[-1]C"
    End With
    On Error GoTo 0
    
    ' Convert the whole range from formulas back to static values
    rng.Value = rng.Value
    
    Application.ScreenUpdating = True
End Sub

Sub FillDown_SpecialCells_MultiCol()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    ' Find the last used row based on Column Q
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    
    ' If the last row is less than 5, there's nothing to process
    If lastRow < 5 Then Exit Sub
    
    ' Set the range to Columns A through H starting from row 5
    Set rng = ws.Range("A5:E" & lastRow)
    
    ' Turn off screen updating and calculations for maximum performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next ' Prevent error if there are no blank cells
    With rng.SpecialCells(xlCellTypeBlanks)
        ' Insert formula to equal the cell exactly one row above
        .FormulaR1C1 = "=R[-1]C"
    End With
    On Error GoTo 0
    
    ' Convert the whole block from formulas back to static values
    rng.Value = rng.Value
    
    ' Turn settings back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub FillPattern_FGH()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim dataArr As Variant
    Dim i As Long, c As Long
    
    Set ws = ActiveSheet
    
    ' Find the last used row based on Column Q
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    
    ' If there's no data to process, exit
    If lastRow < 5 Then Exit Sub
    
    ' Target Columns F, G, and H starting from row 5
    Set rng = ws.Range("F5:H" & lastRow)
    
    ' Read the entire F:H range into an in-memory array (Fastest method)
    dataArr = rng.Value
    
    ' Loop through the array stepping by 3
    ' Array Row 1 = Sheet Row 5 | Array Row 4 = Sheet Row 8 | etc.
    For i = 1 To UBound(dataArr, 1) Step 3
        
        ' Loop through the 3 columns in our array (1=F, 2=G, 3=H)
        For c = 1 To 3
            
            ' If the "master" row in this block has a value...
            If Not IsEmpty(dataArr(i, c)) And dataArr(i, c) <> "" Then
                
                ' Duplicate it to the next 2 rows (checking that we don't exceed the last row)
                If i + 1 <= UBound(dataArr, 1) Then dataArr(i + 1, c) = dataArr(i, c)
                If i + 2 <= UBound(dataArr, 1) Then dataArr(i + 2, c) = dataArr(i, c)
                
            End If
            
        Next c
    Next i
    
    ' Dump the updated array back to the worksheet in one single operation
    rng.Value = dataArr

End Sub

Sub FillDown_SpecialCells_I_to_P()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    
    Set ws = ActiveSheet
    
    ' Find the last used row based on Column Q
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    
    ' If the last row is less than 5, there's nothing to process
    If lastRow < 5 Then Exit Sub
    
    ' Set the range to Columns I through P starting from row 5
    Set rng = ws.Range("I5:P" & lastRow)
    
    ' Turn off screen updating and calculations for maximum performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next ' Prevent error if there are no blank cells
    With rng.SpecialCells(xlCellTypeBlanks)
        ' Insert formula to equal the cell exactly one row above
        .FormulaR1C1 = "=R[-1]C"
    End With
    On Error GoTo 0
    
    ' Convert the whole block from formulas back to static values
    rng.Value = rng.Value
    
    ' Turn settings back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub InsertAndCalculate_ColumnR()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim fData As Variant, rData As Variant
    Dim i As Long, j As Long
    Dim xxDate As Date
    Dim cellParts() As String, itemParts() As String
    Dim totalSum As Double
    Dim rngF As Range
    
    Set ws = ActiveSheet
    
    ' Turn off screen updating and calculation for maximum speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 1. Insert a new Column R (shifts old R to S)
    ws.Columns("R:R").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' 2. Read the reference date 'xx' from the NEW Column S4
    If IsDate(ws.Range("S4").Value) Then
        xxDate = CDate(ws.Range("S4").Value)
    Else
        ' Fallback if it's text but formatted like yyyy/mm/dd
        On Error Resume Next
        xxDate = CDate(ws.Range("S4").Text)
        On Error GoTo 0
    End If
    
    ' 3. Set the new header in R4
    ws.Range("R4").Value = "BeforeColumnR : {" & Format(xxDate, "yyyy/mm/dd") & "}"
    
    ' 4. Find the last row based on Column Q (STATUS)
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 5 Then GoTo CleanExit
    
    ' 5. Load Column F into an in-memory array
    Set rngF = ws.Range("F5:F" & lastRow)
    If rngF.Cells.Count = 1 Then
        ReDim fData(1 To 1, 1 To 1)
        fData(1, 1) = rngF.Value
    Else
        fData = rngF.Value
    End If
    
    ' Prepare an empty array for our Column R results
    ReDim rData(1 To UBound(fData, 1), 1 To 1)
    
    ' 6. Loop through the array in memory, jumping 3 rows at a time (MAXIMUM SPEED)
    For i = 1 To UBound(fData, 1) Step 3
        totalSum = 0 ' Reset sum for the current block
        
        ' Process ONLY the master row of the block (e.g., Row 5, 8, 11...)
        If Not IsEmpty(fData(i, 1)) And fData(i, 1) <> "" Then
            
            ' Split the string by ";"
            cellParts = Split(fData(i, 1), ";")
            
            ' Loop through each date*amount pair
            For j = LBound(cellParts) To UBound(cellParts)
                If Trim(cellParts(j)) <> "" Then
                    
                    ' Split the pair by "*"
                    itemParts = Split(cellParts(j), "*")
                    
                    ' Make sure we actually got a date and an amount
                    If UBound(itemParts) = 1 Then
                        ' Check if the date is strictly before xxDate
                        If IsDate(itemParts(0)) Then
                            If CDate(itemParts(0)) < xxDate Then
                                ' Add to our sum
                                totalSum = totalSum + Val(itemParts(1))
                            End If
                        End If
                    End If
                    
                End If
            Next j
        End If
        
        ' First, set all 3 rows in this current block to 0
        rData(i, 1) = 0
        If i + 1 <= UBound(rData, 1) Then rData(i + 1, 1) = 0
        If i + 2 <= UBound(rData, 1) Then rData(i + 2, 1) = 0
        
        ' Then, overwrite ONLY the row directly below the master (e.g., Row 6, 9, 12...) with totalSum
        If i + 1 <= UBound(rData, 1) Then
            rData(i + 1, 1) = totalSum
        End If
        
    Next i
    
    ' 7. Dump the calculated array into the new Column R in ONE operation
    ws.Range("R5:R" & lastRow).Value = rData

CleanExit:
    ' Restore Excel application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
