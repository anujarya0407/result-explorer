Attribute VB_Name = "Module24"
Sub SnColumnDelete()
Set MR = Range("A1:D1")
    For Each cell In MR
        If cell.Value = "S.No." Then cell.EntireColumn.Delete
    Next
MsgBox "Serial Number Column Deleted."
End Sub

