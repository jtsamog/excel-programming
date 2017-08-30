# word-vba
---
1. [Converting Automatic Numbering to Manual Numbering](https://wordribbon.tips.net/T010248_Converting_Automatic_Numbering_to_Manual_Numbering.html):
  - ALT+F11
  - CTRL+G and input following code into **Immediate**: ActiveDocument.Range.ListFormat.ConvertNumbersToText
  - ENTER
2. [Convert table into text](https://www.extendoffice.com/documents/word/994-word-convert-multiple-tables-to-text.html):
 - ALT+F11
 - Click **Module** on the Insert tab, copy and paste the following VBA code into the Module window;
 - Click **Run** button to apply the VBA
 ```VBA
 Sub TablesToText()
  Dim tbl As Table
  For Each tbl In ActiveDocument.Tables
   tbl.ConvertToText 
   Separator:=wdSeparateByTabs
  Next tbl
  Set tbl = Nothing
 End Sub
 ```
