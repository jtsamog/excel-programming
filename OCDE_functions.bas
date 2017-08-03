Attribute VB_Name = "Module1"
Option Explicit

' ### Calculate "NO. OF WEEKS"
' =ROUNDUP((DATEDIF(G3,H3,"d")/7),0)
Function NumOfWeeks(start_date As Date, end_date As Date)
 
    ' Number of Week - ww
    NumOfWeeks = DateDiff("w", start_date, end_date)
    
    
    ' Number of Day - d
    'NumOfWeeks = DateDiff("d", start_date, end_date) / 7
 
    'MsgBox (ProcessForPaymentDate)
 
End Function


' ### Calculate "PROCESS FOR PAYMENT DATE"
Function ProcessForPaymentDate(start_date As Date, number_of_weeks As Integer)
 ' Get the date half way of total course weeks
 Dim half_course_date As Date
 half_course_date = start_date + (number_of_weeks / 2) * 7
 
 ' if the date is Monday, take the date. Else, get the next Monday from the date.
 If Weekday(half_course_date, 2) = 1 Then
    ProcessForPaymentDate = half_course_date
 Else
    ProcessForPaymentDate = half_course_date + (7 - Weekday(half_course_date, 2) + 1)
 End If
 
 'MsgBox (ProcessForPaymentDate)
 
End Function
