# excel-programming
I can not believe myself coding on excel with VBA. But I have to say that excel programming is so beneficial because:
- No doubt Excel is most popular application for storing data. But excel can not do complex data manipulation. VBA programming solves this shortage.  Therefore, tremendous needs exist for excel programming.
- No need to install anything, all you need is excel. If we code with Python, we have to install Python first. If we use Jave, we have to install JDK first. 
- 

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
### Function
- function return value
- can be called with function name
- can have parameters 
```
'#### Check if a value exist in a range, return rowIndex OR 0
Function isValInRange(ByVal val As String, ByVal rng As Range) As Integer
For Each Row In rng
    If (Row.Cells(, 1) = val) Then
        isValInRange = Row.Row
        Exit Function
    End If
Next
isValInRange = 0
End Function
```
### Sub
- Sub does not return value
- need to call with: Call Sub name
- can have parameters 

## VBA Programming wisdom
- Put your codes into the Module instead of Sheet. 
