Attribute VB_Name = "Module16"
Sub Sno()
Dim MyCol As String
Dim i As Integer
For i = 15 To Range("A" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "S.No.") > 0 Then
Range("A" & i).EntireRow.Delete
End If
Next i
MsgBox "Serial Number Rows Removed"
MsgBox "Step 1 Is Now Completed"
End Sub
