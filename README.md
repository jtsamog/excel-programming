# excel-programming
excel vba programming cases
No doubt Excel is popular. Many people are comfortable using it. So excel VBA programming skills benefits many.

## VBA basic syntax
### Print
MsxBox()
### Variable defination
```
Dim rng As Range
Dim sht As Worksheet
Dim LastRowIndex As Long
Dim SheetName As String
Dim rightMostCol As String
Dim bottomMostRow As Integer

Dim arrStrList() As String
arrStrList = Split("DL1,DL2,DL3,DL4,DL5,DL6,DL7,DL8,DL9", ",")
```
- Dim arrStrList() As String: is Array in VBA;

### Extract values from excel range
```
With Worksheets("data")
bottomMostRow = .Cells(.Rows.Count, "A").End(xlUp).Row
rightMostCol = Left(Replace(.Cells(1, 1).End(xlToRight).Address, "$", ""), 1)
Set rng = .Range("A2:" & rightMostCol & bottomMostRow)

For Each Row In rng.Rows
  Row.Value
  Row.Cells(, 15), Row.Cells(, 9)
Next

.Range("L2:L" & bottomMostRow)

End With
```
- Worksheets("data"): select the sheet named "data"
- bottomMostRow: get the deepest row number
- rightMostCol: get the most right column name;
- Row.Value: get all values in the row
- Row.Cells(, 15): 15th column value in the row
- .Range("L2:L" & bottomMostRow): select certain range
### Date
- Current date:  Date
- Extract Year:  Year(Date)
- Extract Month: DatePart("m", Date)
- Extract Date: DatePart("d", Date)

### Formatting
```
.Range("L1").Value = "Amount to be paid"
.Range("L2:L" & bottomMostRow).NumberFormat = "$#,##0.00"
.Range("L1:M" & bottomMostRow).Font.Bold = True
.Range("L2:M" & bottomMostRow).Font.Color = 16711680
.Range("A2:M" & bottomMostRow).Columns.AutoFit

.Range("I:I").Delete
```

