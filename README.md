 # Rolling Inventory Pipeline: Optimizing Supply Chain Shortage Reports with GAS and VBA Memory Arrays

## Concepts:
- In-Place Backward Loop / Memory Array Method 
- Data Wrangling (or Date String Parsing), HashMapping ,ETL pipeline for multiple tables
- Minimize Time Complexity Approach (For 3 rows as a set reptitivie pattern )
- Consolidated Shortage Table
##  Modules Sequences
1. AddStatusColumnAfterP
2. ExpandQRowsToDemandSupplyBalance
3. FillDown_ItemNo, FillDown_SpecialCells_MultiCol, FillPattern_FGH, FillDown_SpecialCells_I_to_P
4. SupplyWriteIn
5. InsertAndCalculate_ColumnR
6. FillBlanksWithZero
7. CalculateBalance
![Macros Dashboard](macros.png)

### Spreadsheet Size and Range
- Columns: A to AP equals 42 columns.
- Rows: 12,021 rows.
- Total Data Points: 42 columns × 12,021 rows = 504,882 cells.

- Scenario A (Row vs. Row): If your macro loops through all 12,021 rows, and inside that loop, it checks all 12,021 rows again (e.g., to find duplicates or matches), the math is:

```
12,021 × 12,021 = 144,504,441 iterations.
```

- Scenario B (Cell vs. Cell): If you are looping through every single cell and comparing it to every other cell (worst-case scenario):

```
504,882 cells × 504,882 cells = ~254 billion iterations.
```

### 1. Understand / Business Goal  : 
- This project starts from a real supply-chain shortage problem:

The company has a shortage report table with  demand, supply, balance, planned purchase, spot-buy, and vendor response information **spread across different tables**. Manual update is slow because each part number may have multiple vendors, multiple purchase dates, and multiple time buckets.

The business goal is:

> [!TIP]
> **Convert a messy shortage report into a structured, automated planning table** that shows demand, supply, balance, spot-buy, and updated shortage status by date bucket.
:::

- Core business logic:
 **Demand + Supply + SpotBuy = Updated Balance**

- For rolling inventory:
```

Current Balance =
Previous Balance
+ Current Demand
+ Current Supply
+ Current SpotBuy

```
### 2. Match — Mapping Data to Structure

```
Part Number → Target Row
Date → Target Column
Status Type → Target Row Position

```

- The table structure : **Repeated 4-row object per part number.**

Row 1: DEMAND
Row 2: SUPPLY
Row 3: BALANCE
Row 4: SPOTBUY


```c++
// In-memory arrays and stepped itertation loop 
For r = 1 To UBound(shortKeys, 1) Step 3
...
outVals (r+1 , matchCol) =outVals (r+1 ) + qty


PartNo {
Row Index r (e.g., 1): DEMAND

Row Index r + 1 (e.g., 2): SUPPLY

Row Index r + 2 (e.g., 3): BALANCE

Row Index r + 3 (e.g., 4): SPOTBUY

/*
Instead of

1
2
3
4
5
6

iterate

1
4
7
10 
By adding Step 3 in the for loop, 
it loops Iteration 1: r = 1

Iteration 2: r = 4 (1 + 3)

Iteration 3: r = 7 (4 + 3)

Iteration 4: r = 10 (7 + 3)
*/

```
### 3. Plan — Algorithm Design
```
ETL pipeline:

Extract
    Read shortage table
    Read SpotBuy table
    Read Column F planned purchase data
    Read date headers
Transform
    Parse dates
    Parse date*quantity strings
    Match part numbers
    Match date buckets
    Aggregate quantities
    Recalculate balance
Load
    Write result back to shortage report
    Apply formatting
    Highlight abnormal cells

```

### 4. Implement — Data Structures Used

### A. Hash Map / Dictionary Concept
Where used : `SupplyWriteIn(), SpotBuyToShortageReport_Optimized_V4()`
- Part. number -> Shortage Report Row mapping 
- Date -> Timeline Column mapping 
- Hash-map based row and column lookup
    -  to reduce - build : **O(n)**
    -  repeated linear lookup. **O(1)**
- Keywords: Key-Value Store, Hash Table , Stepped Iteration
- Purpose of **Key Iteratio**n : Because Demand , Supply , Balance repeat every three rows.This exploits the dataset’s structure.




- Example:
```c++
// pre-parsing headers and build the O(1) Hash Map 
Dim exactDates As New Collection
exactDates.Add Item:=c, Key:=CStr(CDbl(parsedDate))
/*Meaning:
loops through the timeline Headers (row 4, columns R to AQ) 
Date → Column Number
parse date ubti a double (floating pointer number )and into a string (e.g "46201") as assgined unique key as text identifierm store as a range spreadsheet cell object (c) varible (:= syntax), inserted into VBA collection object exactDateds */ 

keyMap.Add Item:=r + 3, Key:=sKey
/*Meaning:

Part Number → SpotBuy Row

This avoids repeatedly scanning rows and columns.*/
```


- Old logic would be:

For each SpotBuy row
    Search every shortage row
        Search every date column

- New logic:

Build map once
**Use O(1) lookup**

### PseudoCode
```c++
/* 
step 1: Map initialize
step 2: Build memory pointers
step 3: O(1) Lookupp - with numeric value as key that maps to the column index
*/

TargetDate = ParseItemDate(Itemstring)
NumericTarget = ConvertToNumber(TargetDate)

// check if exact date exist in map
IF ExactDatesMap.HashKey (NumericTarget) THEN
    // Retrieval without loop and assign value back
    TargetColumn = ExactDatesMap.GetValue(NumericTarget)
    MemoryGrid[CurrentRow, TargetColumn] = Quantity
ELSE
    // Fallback logic for date ranges
    RunRangeLoopSearch()
END IF
    



```



⸻

## B. Multi-Dimensional Arrays

- Loading  worksheet data into memory:


```c++ 
outVals = wsShort.Range(...).Value
spotData = wsSpot.Range(...).Value
dataArr = ws.Range("R5").Resize(...).Value

/*
Instead of 
Cell
Cell
Cell

This is create up 2D array - 
[data][data][data]
[data][data][data]

This creates 2D arrays:

dataArr(row, column)
inF and outVals are  is  2D Memory array
*/
    Dim inF As Variant
    inF = ws.Range("F5:F" & lastRow).Value
    Dim outVals As Variant
    // FIXED: Exact array bound mapping (No +1 needed)
    outVals = ws.Range("R5").Resize(UBound(inF, 1), colCount).Value
    

```






- Concept:

Worksheet Range → 2D Memory Matrix → Bulk Write Back


- keywords: **Contigious Memory, Memory Array**

### PseudoCode

```c++
/* VBA: 
1. Read into Memory : Grab chunk as data and assigns into variable
Dim arr as Variant 
arr = ws.Range("R5").Resize(lastRow - 4 , colCount).Value
2. Processing in RAM
For r=3 GO UBound (arr,1 ) // inside arr(r,c)
3. Write Back
ws.Range("R5").Resize(UBound(arr,1), colCount).Value =arr
*/

//Step 1: Turn off UI
//step 2: Read data into memory 
// step 3: Process in RAM by looping through the in-memory grid
// Step 4: Write back in one block


// 1
DisableScreenUpdating()
DisableAutomaticCalculations() 
//2 
MemoryGrid = ReadFromSheet (StartCell, EndCell) 
//3 
For rowIndex =3 TO MemoryGrid.LastRow
    Demand = MemoryGrid[rowIndex - 2, Column]
    Supply = MemoryGrid[rowIndex - 1, Column]
    // calculate and update memory
    
    MemoryGrid[rowIndex, Column ] = Demand + Supply
END FOR
//4
WriteToSheet(StartCell, MemoryGrid)
EnableScreenUpdating()
EnableAutomaticCalculations()
```
⸻

## C. Parallel Array for Formatting

`InsertSpotBuyRows()`
- MultiDimential Arrays 

```VBA
outVals(row,column)
countVals(row,column)

/*This tracks how many entries were added to each target cell.
*/

ReDim newData(1 To newTotalRows, 1 to lastColk)

```


-  meaning:

If count > 1, multiple SpotBuy records were merged.


- Concept:

Primary Array: output values
Parallel Array: aggregation count / formatting flag

-  keyword:

metadata array to track aggregation frequency and trigger batch formatting.

```c++
// Step 1: Determine boundaries
OriginalRowCount = SourceData.RowCount
TotalColumns = SourceData.ColumnCount

// Step 2: Calculate the exact size needed for the expanded table
//  add 1 new row for every 3 existing rows
NewRowCount = OriginalRowCount + (OriginalRowCount / 3)

// Step 3: Allocate the exact memory block, empty 2D grid in RAM to prevent memory leaks
AllocateMemoryGrid ExpandedData[Rows: NewRowCount, Columns: TotalColumns]

// Step 4: Populate the new dynamically sized grid
NewRowIndex = 1

FOR OriginalRowIndex = 1 TO OriginalRowCount
    // Copy the existing row into the new grid
    ExpandedData[NewRowIndex] = SourceData[OriginalRowIndex]
    NewRowIndex = NewRowIndex + 1
    
    // Every 3rd row, inject the new custom row (e.g., "SPOTBUY")
    IF OriginalRowIndex MOD 3 == 0 THEN
        ExpandedData[NewRowIndex, TargetColumn] = "SPOTBUY"
        NewRowIndex = NewRowIndex + 1
    END IF
END FOR

// Step 5: Dump the newly allocated and populated grid back to the sheet
WriteToSheet(StartCell, ExpandedData)

```

## D. Syntax Parsing 
```c++
strF = Trim(CStr(inF(r,1))) 

if Len(strF)>0 Then 
    parts = Split(strF, ";")
    For i = LBound(parts) To UBound(parts)
        itemParts = Split(Trim(parts(i)), "*")


/*
[Column F String] ---> "2026/07/02*117;2026/07/09*140"
                             |
                      (Split by ";")
                             |
               +-------------+-------------+
               |                           |
     "2026/07/02*117"            "2026/07/09*140"
               |                           |
         (Split by "*")              (Split by "*")
               |                           |
     Date: 2026/07/02            Date: 2026/07/09
     Qty: 117                    Qty: 140
               |                           |
   [exactDates Hash Lookup]     [exactDates Hash Lookup]
               |                           |
       Maps to Col: 1              Maps to Col: 2
               |                           |
               v                           v
  outVals(r+1, 1) = 117       outVals(r+1, 2) = 140
  (Placed directly into the SUPPLY rows inside RAM) */
```
⸻
## Evaluation- Time Complexity 

1. Read the entire range of data into a VBA Variant Array in memory (1 interaction).

2. Loop through that array in memory to build a new array with the extra rows inserted. Memory operations are virtually instantaneous.

3. Write the new array back to the worksheet in one go (1 interaction).

- Rolling State  : 

```
Balance (i ) depends on Balance (i=1)
```
- Original **O(n x m xk)**
```
Rows -> columns -> Headers
```

- New **O (n x m)**

```
Preprocessed Headers -> Rows -> O(1) lookup
```
with precomputes 
```
Header Dates , Heade Range , HashMaps
```



---
## Shortage Gap Calculation
In a typical supply chain or MRP (Material Requirements Planning) shortage report, **negative values** represent a projected deficit where demand exceeds your available supply.

The relationship between the **-292** in the first week (07/08~07/14) and the **-100** in the second week (07/15~07/21) depends entirely on how your report calculates its totals over time. There are two common ways this is handled:

### 1. Cumulative Projected Balance (Running Total)

If the report displays a running total of your inventory balance, the relationship is sequential and cumulative.

* **The Relationship:** The shortage is actually improving. You end the first week short 292 units, but by the end of the second week, your shortage drops to 100 units.
* **What it means:** You are likely expecting a scheduled receipt (e.g., a Purchase Order or Work Order delivery) of exactly 192 units during the week of 07/15~07/21. This new supply partially offsets your previous shortage:

$$-292 + 192 = -100$$

![rollingtotal](rollingtotal.png)
![shortagereport](shortagereport.png)


### 2. Discrete Net Requirements (Period-Specific)

If the report shows isolated "net requirements" per time bucket, the values represent independent new shortages occurring only within those specific date ranges.

* **The Relationship:** The values are additive. You have independent gaps in supply for both weeks.
* **What it means:** You are short 292 units to fulfill demand in the first week, and you are short an *additional* 100 units to fulfill new demand in the second week. To resolve both, you would need to expedite or order a total of **392 units**.

---

### 🔍 How to tell which report you are looking at?

Look at the **"Total Shortage"** column for that specific part number:

| If the Total Shortage is... | Then the report type is... |
| --- | --- |
| **-100** (matches the final bucket) | 📈 **Cumulative Report** |
| **-392** (or greater, if there are other shortages) | 📊 **Discrete / Period-Specific Report** |


## Stock Arrives Calculation for Discrete Net Requirements 


In a **Discrete Net Requirements** report, numbers are isolated into specific time buckets. Because of this, how newly purchased stock impacts your report depends entirely on **when the stock arrives** (the scheduled delivery date).

Using your example, here is how the numbers change based on arrival timing:

### 🚀 Scenario A: Stock arrives exactly when needed (Week 2)

If you buy **100 units** scheduled to arrive between **07/15 ~ 07/21**:

* **07/08 ~ 07/14:** Remains `-292` *(the new stock hasn't arrived yet).*
* **07/15 ~ 07/21:** Changes from `-100` to `0` *(the shortage for this specific week is fully covered).*

### ⏱️ Scenario B: Stock arrives early (Week 1)

If you buy **100 units** and they arrive early between **07/08 ~ 07/14**:

* **07/08 ~ 07/14:** Changes from `-292` to `-192` *(the shortage is partially covered).*
* **07/15 ~ 07/21:** Remains `-100` *(this remains a separate, independent shortage for the second week).*

### ⚠️ Scenario C: Stock arrives late (Week 2)

If you buy **400 units** to try and cover both shortages, but they do not arrive until **07/15 ~ 07/21**:

* **07/08 ~ 07/14:** Remains `-292` *(you missed the deadline for the first week's demand).*
* **07/15 ~ 07/21:** Changes from `-100` to `+300` *(you cover Week 2 and leave a surplus, but the line starved the week prior).*

> 💡 **Key Takeaway:** In a discrete report, late stock cannot retroactively fix a past shortage. Timing is everything.

---

### 📊 Summary Impact Matrix

| Scenario | Qty Added | Arrival Window | Week 1 Status (07/08~07/14) | Week 2 Status (07/15~07/21) |
| --- | --- | --- | --- | --- |
| **Baseline Shortage** | — | — | `-292` | `-100` |
| **A (Just-in-Time)** | +100 | Week 2 | `-292` | `0` (Resolved) |
| **B (Early)** | +100 | Week 1 | `-192` (Partial) | `-100` |
| **C (Late)** | +400 | Week 2 | `-292` (Missed) | `+300` (Surplus) |

### Aim:
- Batch Q:AP
- Single Active Cell Logging
- Column F Parsing 
- Daily / Weekly / Overhead Matching 
- Adjusted-Current Calculation 

### Overall Flow

For the first shortage value:
- Begining Stock + Column F Qty before/current period + xValue = First Shortage Balance
For the subsequent shortage value 
- Previous Balance + Column F Qty between previous shortage and current shortage + xValue = Current Shortage Balance
- xValue = Current Shortage Balance - Previous Balance - Column F Qty 

Begining balance = 10
AE Column F quantity = 90
Balance before AF = 10 + 90 = 100
AF current value = -90
Movement = -90 - 100 = -190

```
Read:
    Current Inventory Value / Balance for each date range
    +
    Previous Inventory Values
    +
    Initial Stock ( L + M+ N)
    +
    Column F planned quantities
    
↓
Calculate Movement (xValue)
↓
Write one record into Movements:
(Date, Item No, Movement)

Current = Current + Column F Qty
Movement = Current - Previous

    
```

```javascript
/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */
/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */


/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */



## VBA Usage Setup Instrutions



### Windows (Excel Desktop)

1. Open the attached `.xlsm` file.
2. If a security warning appears at the top of the screen, click **Enable Editing**, and then click **Enable Content**.
3. Press **Alt + F8** to open the Macro list.
4. Run either of the following macros:
   * **LogMovement**: Calculates the currently selected single cell.
   * **LogAllMovements**: Batch calculates all rows (starting from Row 5) for columns Q to AP, and writes the results into the `Movements` worksheet.

---

### Mac (Excel Desktop)

1. Open the `.xlsm` file.
2. If the system prompts you to enable macros, select **Enable Macros**.
3. Click **Tools** → **Macro** → **Macros** (or **Developer** → **Macros**, depending on your Excel version).
4. Run **LogMovement** or **LogAllMovements**.

---

> [!IMPORTANT]
> **Reminders:**
> * Please use Microsoft Excel Desktop to open the file; the web version of Excel cannot run VBA macros.
> * Please keep the file in `.xlsm` format. If you save it as `.xlsx`, the macro code will be removed.
