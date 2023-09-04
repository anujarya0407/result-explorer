Attribute VB_Name = "Module8"
Sub myDeleteRows2()
Dim MyCol As String
Dim i As Integer
For i = 1 To Range("A" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "TOTAL(GRADE)") > 0 Or Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "A: Absent") > 0 Or Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "D: Detained") > 0 Then
Range("A" & i).EntireRow.Delete
End If
Next i
MsgBox "Total(Grade),A:Absent and D:Detained Rows Removed."
End Sub
