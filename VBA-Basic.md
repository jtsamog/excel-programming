# VBA-BASIC
---
### Conditional Statement
#### IF ELSE statement
```VBA
If (credit = 0 And grade_code = "P") Then
    sumIfAmount = 30
ElseIf (credit = 3 And grade_code = "W") Then
    'sumIfAmount = 200
ElseIf (credit = 3) Then
    sumIfAmount = 400
End If
```

### What NOT Works
1. GoTo statement in IF ELSE not work.
