Attribute VB_Name = "Module9"
Sub myDeleteRows3()
Dim MyCol As String
Dim i As Integer
For i = 1 To Range("A" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "CS: Credits Secured") > 0 Or Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "AP: Already Passed") > 0 Or Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "Result of Programme Code: 020") > 0 Then
Range("A" & i).EntireRow.Delete
End If
Next i
MsgBox "CS:Credit Secured, AP:Already Passed and Result of Programme Rows Removed""
End Sub
