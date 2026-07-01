# ConsolidatedShortageTable-GoogleAppsScript

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
