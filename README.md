# ConsolidatedShortageTable-GoogleApps/VBA
## Aim:
- Fetch Batch Q:AP
- Single Active Cell Logging
- Column F Parsing 
- Daily / Weekly / Overhead Matching 
- Adjusted-Current Calculation 

### Overall Flow
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
## UPDATED Flow and Functions
```
Subroutines (Sub)
AddStatusColumnAfterP()
ExpandQRowsToDemandSupplyBalance()
SupplyWriteIn()
FillBlanksWithZero()
CalculateBalances()
HighlightNegativeBalances()
FillDown_ItemNo()
InsertAndCalculate_ColumnR()
Helper Functions (Function)
ParseMD(txt) - converts MM/DD
ParseYMD(txt) - YYYY/MM/DD

```


## Performance Optimizations (Reducing Time Complexity)

To handle massive datasets efficiently, the new VBA scripts were completely refactored to prioritize in-memory processing and bulk worksheet operations.

### 1. $O(1)$ Header Evaluation vs. Nested Loops
* **The Old Way:** Date headers were evaluated inside a nested loop for every row and then every column. For 1,000 rows and 26 columns, the script was forced to evaluate string parsing rules 26,000 times.
* **The New Way:** Headers are read, precalculated, and stored into variables in memory **once**. The script then compares the parsed column directly against the memory array, drastically reducing execution time.

### 2. In-Memory Array Processing
* **Found in:** `CalculateBalances`, `FillPattern_FGH`, `InsertAndCalculate_ColumnR`
* **Mechanism:** Reading and writing to individual worksheet cells is the slowest operation in VBA. Data is now sucked into memory blocks instantly using `arr = ws.Range(...).Value`, calculated at maximum CPU speed, and dumped back to the sheet in a single action.

### 3. Stepped Looping (Jumping 3-Row Blocks)
* **Found in:** `CalculateBalances`, `FillPattern_FGH`, `InsertAndCalculate_ColumnR`
* **Mechanism:** Because the data is strictly grouped into 3-row blocks (*Demand, Supply, Balance*), loops now use `For i = 1 To UBound(data) Step 3`. This skips 2 out of every 3 rows. If the sheet has 30,000 rows, the code only performs 10,000 iterations.

### 4. Bulk Object Manipulation ($O(1)$ Worksheet I/O)
* **Filling Blanks:** (Found in `FillBlanksWithZero`, `FillDown_ItemNo`, `FillDown_SpecialCells_MultiCol`) Instead of iterating `If cell.Value = "" Then`, the scripts leverage Excel's native C++ backend via `.SpecialCells(xlCellTypeBlanks)` to fill all blank cells instantly.
* **Formatting:** (Found in `HighlightNegativeBalances`) Instead of looping through rows to check if a number is negative and changing the `.Font.Color` one cell at a time, the code applies a **Conditional Formatting Rule** to the entire grid (`R7:AR[lastRow]`) at once. Excel is given the mathematical logic once:
  ```excel
  =AND(R7<0, MOD(ROW(),3)=1)


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
