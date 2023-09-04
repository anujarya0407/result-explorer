Attribute VB_Name = "Module7"
Sub myDeleteRows1()
Dim MyCol As String
Dim i As Integer
For i = 1 To Range("A" & "65536").End(xlUp).Row Step 1
If Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "LEGEND") > 0 Or Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "Internal") > 0 Or Application.WorksheetFunction.CountIf(Range("A" & i & ":AZ" & i), "PAPERID(CREDITS)") > 0 Then
Range("A" & i).EntireRow.Delete
End If
Next i

MsgBox "Legend,Internal,PaperID rows removed."
End Sub
