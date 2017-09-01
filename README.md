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

## Create a custom function
 - create the functin in the Module
 - Apply it with =function  same as using default functions
 - can save as Add-ons (xlam file) for later use
 - https://support.office.com/en-us/article/Create-Custom-Functions-in-Excel-2007-2f06c10b-3622-40d6-a1b2-b6748ae8231f
## Check duplication
* Highlight row(s) of duplicates
* Create Duplicate tab with nothing automatically going to it
* Evaluate columns: SATURN_SPRIDEN_SPRIDEN_ID [A], SSBSECT_CRN[F], SFRSTCR_GRDE_CODE[I]
```
With Worksheets("Self-paced")

'Declaring the lastRow variable as Long to store the last row value in the Column L
Dim lastRow As Long
'matchFoundIndex is to store the match index values of the given value
Dim matchFoundIndex As Long
'rowid is to loop through all the records in the column 1 using For loop
Dim rowid As Long
'Finding the last row in the Column 1
lastRow = .Cells(.Rows.Count, "L").End(xlUp).Row
' Convert Range to Array() in order to use Match function
Dim arrStrList2() As String
Dim str As String
Dim rng33 As Range
i = 0
ReDim arrStrList2(lastRow - 1)
Set rng33 = .Range("A1:P" & lastRow)
For Each Row In rng33.Rows
'MsgBox (Row.Cells(, 1) & Row.Cells(, 2))
arrStrList2(i) = Row.Cells(, 1) & Row.Cells(, 6) & Row.Cells(, 9)
i = i + 1
Next

For rowid = 1 To lastRow
    'getting match index number for the value of the cells = [A,F,I] === [1,6,9]
    str = .Cells(rowid, 1) & .Cells(rowid, 6) & .Cells(rowid, 9)
    matchFoundIndex = WorksheetFunction.Match(str, arrStrList2, 0)
    If rowid <> matchFoundIndex Then
        'Highlight the row in RED color
        .Range("A" & rowid & ":P" & rowid).Interior.ColorIndex = 3 ' RED Color
    End If

Next

End With
```
