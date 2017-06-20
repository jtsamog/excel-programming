# excel-programming
excel vba programming cases
No doubt Excel is popular. Many people are comfortable using it. So excel VBA programming skills benefits many.

## VBA basic syntax
### Print
MsxBox()
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

End With
```
- Worksheets("data"): select the sheet named "data"
- bottomMostRow: get the deepest row number
- rightMostCol: get the most right column name;
- Row.Value: get all values in the row
- Row.Cells(, 15): 15th column value in the row
### Date
- Current date:  Date
- Extract Year:  Year(Date)
- Extract Month: DatePart("m", Date)
- Extract Date: DatePart("d", Date)

