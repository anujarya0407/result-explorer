Attribute VB_Name = "Module10"
Sub myDeleteRows4()
Dim MyCol As String
Dim i As Integer
For i = 1 To Range("A" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "RESULT TABULATION SHEET") > 0 Then
Range("A" & i).EntireRow.Delete
End If
Next i
MsgBox "Result Tabulation Sheet Rows Removed"
End Sub
