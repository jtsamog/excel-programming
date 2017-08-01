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

### Soring
- Up to 3 columns: Worksheets(sheetName).Range("A1:D10").Sort key1:=.Range("B1:B10"), order1:=xlAscending, Header:=xlYes
- More than 3, consider followings:
```
Dim wks As Worksheet
Set wks = Worksheets("for Email")
bottomMostRow = wks.Cells(wks.Rows.Count, "H").End(xlUp).Row
With wks.Sort
With .SortFields
.Add Key:=wks.Range("G1:G" & bottomMostRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
.Add Key:=wks.Range("H1:H" & bottomMostRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
.Add Key:=wks.Range("C1:C" & bottomMostRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
.Add Key:=wks.Range("D1:D" & bottomMostRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
.Add Key:=wks.Range("A1:A" & bottomMostRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
End With
.SetRange Range("A1:M" & bottomMostRow)
.Header = xlYes
.MatchCase = False
.Orientation = xlTopToBottom
'.SortMethod = xlPinYin
.Apply
End With
```

### Conver string to formula
- https://www.extendoffice.com/documents/excel/1683-excel-convert-text-string-to-formula.html

### Set background color
- [color index](http://dmcritchie.mvps.org/excel/colors.htm)
- Worksheets.Range("Q3").Interior.ColorIndex = 6
- Worksheets.Range("Q3").Interior.Color = RGB(255,255,0)
## VBA Programming wisdom
- Put your codes into the Module instead of Sheet. 
## Excel Formula
- Count weeks: =ROUNDUP((DATEDIF(date1,date2,"d")/7),0)
- First Monday after a date: =date+(7-WEEKDAY(date,2)+1)
- IF Monday, set the current value, ELSE, next Monday: =IF(WEEKDAY(date)=2,date,date+(7-WEEKDAY(date,2)+1))
