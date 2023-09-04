Attribute VB_Name = "Module6"
Sub SchemeID()
Dim MyCol As String
Dim i As Integer
For i = 1 To Range("C" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "SchemeID: 480202015001") > 0 Then
Range("A" & i).EntireRow.Delete
End If
Next i
MsgBox "SchemeID Removed"
End Sub
