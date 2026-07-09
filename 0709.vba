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

Sub InsertSpotBuyRows()
    Dim ws As Worksheet
    Dim sourceData As Variant, newData As Variant
    Dim lastRow As Long, lastCol As Long
    Dim totalRows As Long, newTotalRows As Long
    Dim i As Long, j As Long, newRowIdx As Long
    
    ' Set to your active sheet
    Set ws = ActiveSheet
    
    ' Find the last row based on Column P (Total Shortage) to avoid grabbing unrelated data below
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    If lastRow < 5 Then Exit Sub ' No data to process
    
    ' Dynamically find the last column based on Row 4 (Headers)
    ' This ensures Columns R, S, T, etc., are all captured up to Column AS (or beyond)
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    
    ' 1. Read the ENTIRE table (A5 to LastRow, LastCol) into memory
    sourceData = ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol)).Value
    totalRows = UBound(sourceData, 1)
    
    ' Calculate the size of the new array (adding 1 row for every 3 existing rows)
    newTotalRows = totalRows + Int(totalRows / 3)
    
    ' Size the new array to hold the expanded data, across ALL columns
    ReDim newData(1 To newTotalRows, 1 To lastCol)
    
    newRowIdx = 1
    
    ' 2. Process data entirely in memory
    For i = 1 To totalRows
        ' Copy the current row's data (All columns) to the new array
        For j = 1 To lastCol
            newData(newRowIdx, j) = sourceData(i, j)
        Next j
        newRowIdx = newRowIdx + 1
        
        ' Every 3rd row (after DEMAND, SUPPLY, BALANCE), inject the "SPOTBUY" row
        If i Mod 3 = 0 Then
            ' Copy Columns A to P (1 to 16) from the row above
            For j = 1 To 16
                newData(newRowIdx, j) = sourceData(i, j)
            Next j
            
            ' Set Column Q (17) to "SPOTBUY"
            newData(newRowIdx, 17) = "SPOTBUY"
            
            ' For Columns R onwards (18 to lastCol), the new row will automatically be blank.
            ' If you want them to be exactly 0 instead of blank, uncomment the following 3 lines:
            ' For j = 18 To lastCol
            '     newData(newRowIdx, j) = 0
            ' Next j
            
            newRowIdx = newRowIdx + 1
        End If
    Next i
    
    ' 3. Write the entire expanded array back to the worksheet
    ws.Range("A5").Resize(newTotalRows, lastCol).Value = newData
    
    MsgBox "Processing complete!", vbInformation
End Sub

Sub SpotBuyToShortageReport_Optimized_V4()
    Dim wsSpot As Worksheet, wsShort As Worksheet
    Dim lastRowSpot As Long, lastRowShort As Long, lastColShort As Long
    
    ' Set your worksheets (adjust names as necessary)
    Set wsSpot = ThisWorkbook.Sheets("SpotBuy")
    Set wsShort = ThisWorkbook.Sheets("TestFile2 (2)")
    
    ' --- CONFIGURATION ---
    Const colSpotKey As Long = 2    ' Column B in SpotBuy
    Const colSpotData As Long = 20  ' Column T in SpotBuy (Contains Date*Qty)
    Const colShortKey As Long = 1   ' Column A in Shortage Report
    Const startColShort As Long = 18 ' Column R in Shortage Report
    Const startRow As Long = 5
    
    ' Get Last Rows
    lastRowSpot = wsSpot.Cells(wsSpot.Rows.Count, colSpotKey).End(xlUp).Row
    lastRowShort = wsShort.Cells(wsShort.Rows.Count, colShortKey).End(xlUp).Row
    lastColShort = wsShort.Cells(4, wsShort.Columns.Count).End(xlToLeft).Column
    
    If lastRowSpot < 5 Or lastRowShort < startRow Then Exit Sub
    
    Dim c As Long, r As Long
    
' ==========================================
    ' STEP 1: PARSE HEADERS & DATES (O(N))
    ' ==========================================
    Dim colCount As Long
    colCount = lastColShort - startColShort + 1
    
    Dim exactDates As New Collection
    
    ' Variables for the boundaries
    Dim beforeDateLimit As Date, overDateLimit As Date
    Dim beforeDateCol As Long, overDateCol As Long
    
    Dim hVal As String
    Dim headerVals As Variant
    Dim parsedDate As Date
    
    ' Read headers R4 to end
    headerVals = wsShort.Range(wsShort.Cells(4, startColShort), wsShort.Cells(4, lastColShort)).Value
    
    For c = 1 To colCount
        hVal = Trim(CStr(headerVals(1, c)))
        
        ' 1. Check for BEFORE format: e.g., "{2026/07/07}"
        If InStr(hVal, "{") > 0 And InStr(hVal, "}") > 0 Then
            Dim tempDateStr As String
            tempDateStr = Split(Split(hVal, "{")(1), "}")(0)
            beforeDateLimit = ParseYMD(tempDateStr)
            beforeDateCol = c
            
        ' 2. Check for OVER format: e.g., "Over10/05"
        ElseIf Left(hVal, 4) = "Over" Then
            Dim mdStr As String
            mdStr = Replace(hVal, "Over", "")
            overDateLimit = ParseMD_Year(mdStr, 2026) ' Assumes year 2026, adjust if needed
            overDateCol = c
            
        ' 3. RANGE FORMAT: e.g., "07/21~07/27"
        ElseIf InStr(hVal, "~") > 0 Then
            Dim rangeStart As String, rangeEnd As String
            Dim dStart As Date, dEnd As Date
            Dim d As Date
            
            rangeStart = Split(hVal, "~")(0) ' Extracts "07/21"
            rangeEnd = Split(hVal, "~")(1)   ' Extracts "07/27"
            
            dStart = ParseMD_Year(rangeStart, 2026)
            dEnd = ParseMD_Year(rangeEnd, 2026)
            
            ' Loop through EVERY day in the range and assign it to this column!
            If dStart > 0 And dEnd >= dStart Then
                For d = dStart To dEnd
                    On Error Resume Next
                    exactDates.Add Item:=c, Key:=CStr(CDbl(d))
                    On Error GoTo 0
                Next d
            End If
            
        ' 4. Standard Date Mapping
        Else
            parsedDate = ParseYMD(hVal)
            If parsedDate > 0 Then
                On Error Resume Next ' Collection keys must be unique strings
                exactDates.Add Item:=c, Key:=CStr(CDbl(parsedDate))
                On Error GoTo 0
            End If
        End If
    Next c
    ' ==========================================
    ' STEP 2: MAP SHORTAGE REPORT KEYS (O(N))
    ' ==========================================
    Dim keyMap As New Collection
    Dim shortKeys As Variant
    shortKeys = wsShort.Range(wsShort.Cells(startRow, colShortKey), wsShort.Cells(lastRowShort, colShortKey)).Value
    
    ' Loop through Shortage Report keys (Step 4 for Demand/Supply/Balance/SpotBuy sets)
    For r = 1 To UBound(shortKeys, 1) Step 4
        Dim sKey As String
        sKey = Trim(CStr(shortKeys(r, 1)))
        If Len(sKey) > 0 Then
            On Error Resume Next
            ' Store the Row Index for the 4th row (SpotBuy) of this set
            keyMap.Add Item:=r + 3, Key:=sKey
            On Error GoTo 0
        End If
    Next r
    
    ' ==========================================
    ' STEP 3: PROCESS SPOTBUY DATA IN MEMORY
    ' ==========================================
    Dim spotData As Variant
    ' Read SpotBuy data into memory (up to Column T)
    spotData = wsSpot.Range(wsSpot.Cells(startRow, 1), wsSpot.Cells(lastRowSpot, colSpotData)).Value
    
    ' Prepare Output Arrays for Shortage Report (Values and Sum Counts)
    Dim outVals As Variant
    Dim countVals() As Long ' Tracks how many times a cell was added to (for Red/Bold)
    outVals = wsShort.Range(wsShort.Cells(startRow, startColShort), wsShort.Cells(lastRowShort, lastColShort)).Value
    ReDim countVals(1 To UBound(outVals, 1), 1 To UBound(outVals, 2))
    
    Dim tRow As Long, tCol As Long
    Dim eDate As Date, qty As Double
    Dim currentKey As String
    Dim cellData As String
    Dim dataParts() As String, itemParts() As String
    Dim pIdx As Long
    
    For r = 1 To UBound(spotData, 1)
        currentKey = Trim(CStr(spotData(r, colSpotKey)))
        
        If Len(currentKey) > 0 Then
            tRow = 0
            On Error Resume Next
            tRow = keyMap(currentKey)
            On Error GoTo 0
            
            If tRow > 0 Then
                cellData = Trim(CStr(spotData(r, colSpotData)))
                If Len(cellData) > 0 Then
                    dataParts = Split(cellData, ";")
                    
                    For pIdx = LBound(dataParts) To UBound(dataParts)
                        itemParts = Split(Trim(dataParts(pIdx)), "*")
                        
                        If UBound(itemParts) = 1 Then
                            eDate = ParseYMD(itemParts(0))
                            qty = CDbl(Val(itemParts(1)))
                            
                            If qty <> 0 And eDate > 0 Then
                                tCol = 0
                                
                                ' === THE OPTIMIZED CONDITIONAL CHECKS ===
                                If beforeDateLimit > 0 And eDate < beforeDateLimit Then
                                    tCol = beforeDateCol ' Assign to Before limit column
                                    
                                ElseIf overDateLimit > 0 And eDate > overDateLimit Then
                                    tCol = overDateCol ' Assign to Over limit column
                                    
                                Else
                                    ' O(1) Exact Date Lookup
                                    On Error Resume Next
                                    tCol = exactDates(CStr(CDbl(eDate)))
                                    On Error GoTo 0
                                End If
                                ' ========================================
                                
                                ' Accumulate Values
                                If tCol > 0 Then
                                    If IsEmpty(outVals(tRow, tCol)) Or outVals(tRow, tCol) = "" Then
                                        outVals(tRow, tCol) = qty
                                    Else
                                        outVals(tRow, tCol) = outVals(tRow, tCol) + qty
                                    End If
                                    countVals(tRow, tCol) = countVals(tRow, tCol) + 1
                                End If
                            End If
                        End If
                    Next pIdx
                End If
            End If
        End If
    Next r
    
    ' ==========================================
    ' STEP 4: WRITE BACK & APPLY FORMATTING
    ' ==========================================
    Application.ScreenUpdating = False
    
    Dim destRange As Range
    Set destRange = wsShort.Range(wsShort.Cells(startRow, startColShort), wsShort.Cells(lastRowShort, lastColShort))
    destRange.Value = outVals
    
    ' V3 FIX: Clear formatting ONLY on SpotBuy rows (every 4th row)
    Dim clearRange As Range
    For r = 4 To UBound(outVals, 1) Step 4
        If clearRange Is Nothing Then
            Set clearRange = destRange.Rows(r)
        Else
            Set clearRange = Union(clearRange, destRange.Rows(r))
        End If
    Next r
    
    If Not clearRange Is Nothing Then
        clearRange.Font.Bold = False
        clearRange.Font.ColorIndex = xlAutomatic
    End If
    
    ' Batch formatting using Union (Extremely fast)
    Dim formatRange As Range
    For r = 1 To UBound(countVals, 1)
        For c = 1 To UBound(countVals, 2)
            If countVals(r, c) > 1 Then
                If formatRange Is Nothing Then
                    Set formatRange = destRange.Cells(r, c)
                Else
                    Set formatRange = Union(formatRange, destRange.Cells(r, c))
                End If
            End If
        Next c
    Next r
    
    If Not formatRange Is Nothing Then
        formatRange.Font.Bold = True
        formatRange.Font.Color = vbRed
    End If
    
    Application.ScreenUpdating = True
    MsgBox "SpotBuy mapping and summation complete!", vbInformation
End Sub


' NEW: Parses MM/DD headers and attaches a specific year
Function ParseMD_Year(ByVal txt As String, ByVal yr As Long) As Date
    On Error Resume Next
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, vbCr, "")
    Dim p() As String
    p = Split(Trim(txt), "/")
    If UBound(p) = 1 Then
        ParseMD_Year = DateSerial(yr, CLng(p(0)), CLng(p(1)))
    End If
    On Error GoTo 0
End Function

Sub FillSpotBuyBlanksWithZero()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colCount As Long
    Dim dataArr As Variant
    Dim r As Long, c As Long
    
    Set ws = ActiveSheet
    
    ' Find the last row based on Column Q (STATUS)
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 8 Then Exit Sub ' Nothing to process if we haven't reached the first SpotBuy row
    
    ' Columns R (18) to AR (44) equals 27 columns
    colCount = 44 - 18 + 1
    
    ' --- STEP 1: LOAD TO MEMORY (Maximum Speed) ---
    dataArr = ws.Range("R5").Resize(lastRow - 4, colCount).Value
    
    ' --- STEP 2: PROCESS IN MEMORY ---
    ' Start at Array Index 4 (Row 8) and jump 4 rows at a time
    For r = 4 To UBound(dataArr, 1) Step 4
        For c = 1 To colCount
            ' Check if the cell is truly empty or just contains a blank string
            If IsEmpty(dataArr(r, c)) Or Trim(CStr(dataArr(r, c))) = "" Then
                dataArr(r, c) = 0
            End If
        Next c
    Next r
    
    ' --- STEP 3: WRITE BACK IN ONE GO ---
    ws.Range("R5").Resize(UBound(dataArr, 1), colCount).Value = dataArr
    
    MsgBox "SpotBuy blanks filled with 0 successfully!", vbInformation
End Sub

Sub AddSpotBuyToBalance_RollingRecalc()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colCount As Long
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim prevBal As Double, currDemand As Double, currSupply As Double, currSpotBuy As Double
    
    Set ws = ActiveSheet
    
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 8 Then Exit Sub
    
    colCount = 44 - 18 + 1 ' Columns R to AR
    dataArr = ws.Range("R5").Resize(lastRow - 4, colCount).Value
    
    ' Start at Array Index 4 (Row 8) and jump 4 rows at a time
    For r = 4 To UBound(dataArr, 1) Step 4
        
        ' --- 1. Calculate the First Column (Column R / c = 1) ---
        ' R7 = original R7 + R8
        dataArr(r - 1, 1) = Val(CStr(dataArr(r - 1, 1))) + Val(CStr(dataArr(r, 1)))
        
        ' --- 2. Calculate the Rolling Columns (Columns S to AR / c = 2 to colCount) ---
        For c = 2 To colCount
            ' Grab all 4 variables needed for a true rolling inventory balance
            prevBal = Val(CStr(dataArr(r - 1, c - 1)))     ' The newly updated Balance from the left
            currDemand = Val(CStr(dataArr(r - 3, c)))      ' Current column's Demand (e.g., Row 5)
            currSupply = Val(CStr(dataArr(r - 2, c)))      ' Current column's Supply (e.g., Row 6)
            currSpotBuy = Val(CStr(dataArr(r, c)))         ' Current column's SpotBuy (e.g., Row 8)
            
            ' S7 = R7(New) + S5(Demand) + S6(Supply) + S8(SpotBuy)
            dataArr(r - 1, c) = prevBal + currDemand + currSupply + currSpotBuy
        Next c
    Next r
    
    ws.Range("R5").Resize(UBound(dataArr, 1), colCount).Value = dataArr
    MsgBox "SpotBuy added and rolling balances accurately recalculated!", vbInformation
End Sub
