Attribute VB_Name = "Module1"
Option Explicit

' ################# Calculate "NO. OF WEEKS" ###################################################
' Did not use this function, used following Formula instead
' = ROUNDUP((DATEDIF(G3,H3,"d")/7),0)
Function NumOfWeeks(start_date As Date, end_date As Date)
    ' Number of Week - ww
    NumOfWeeks = DateDiff("w", start_date, end_date)
    
    ' Number of Day - d
    'NumOfWeeks = DateDiff("d", start_date, end_date) / 7
    'MsgBox (ProcessForPaymentDate)
End Function


' ################# Calculate "PROCESS FOR PAYMENT DATE" ########################################
' For 6-12 weeks, processing is on Monday following completion of the 4th week
' Treat a 15 week course as 16 because I don’t think we have 15 week courses
' 20 weeks, the Monday following completion of 10th week
' 24-26 weeks, the Monday following completion of the 12th week
Function ProcessForPaymentDate(start_date As Date, number_of_weeks As Integer)
    ' Varible for deciding process weeks
    Dim process_weeks As Integer
    ' For 6-12 weeks, processing is on Monday following completion of the 4th week
    If (number_of_weeks >= 0 And number_of_weeks <= 12) Then
        process_weeks = 4
    ' Treat a 15 week course as 16 because I don’t think we have 15 week courses
    ElseIf (number_of_weeks = 15 Or number_of_weeks = 16) Then
        process_weeks = 8
    ' 20 weeks, the Monday following completion of 10th week
    ElseIf (number_of_weeks = 20) Then
        process_weeks = 10
    ' 24-26 weeks, the Monday following completion of the 12th week
    ElseIf (number_of_weeks >= 24 And number_of_weeks <= 26) Then
        process_weeks = 12
    Else
        process_weeks = 0
    End If
    
    ' Varible for the process date
    Dim process_date As Date
    process_date = start_date + process_weeks * 7
    
    ' If process_weeks are not covered, point it out
    If process_date = start_date Then
        ProcessForPaymentDate = "NOT RIGHT"
    ' If the date is Monday, take the date. Else, get the next Monday from that date.
    ElseIf Weekday(process_date, 2) = 1 Then
        ProcessForPaymentDate = process_date
    Else
        ProcessForPaymentDate = process_date + (7 - Weekday(process_date, 2) + 1)
    End If
 
    'MsgBox (ProcessForPaymentDate)
 
End Function
