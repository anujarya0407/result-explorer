Attribute VB_Name = "Module23"
Sub SubjectCode()
Dim MyCol As String
Dim I As Integer
For I = 2 To Range("B" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("B" & I & ":BZ" & I), "20101(4)") > 0 Then
Range("B" & I).EntireRow.Delete
End If
Next I
MsgBox "Subject Codes Removed"
End Sub

