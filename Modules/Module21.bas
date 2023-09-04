Attribute VB_Name = "Module21"
Sub PhotoColumnDelete()
'updateby Extendoffice 20160616
    Dim xEndCol As Long
    Dim I As Long
    Dim xDel As Boolean
    On Error Resume Next
    xEndCol = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If xEndCol = 0 Then
        MsgBox "There is no data on """ & ActiveSheet.Name & """ .", vbExclamation, "Kutools for Excel"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For I = xEndCol To 1 Step -1
        If Application.WorksheetFunction.CountA(Columns(I)) <= 1 Then
            Columns(I).Delete
            xDel = True
        End If
    Next
    If xDel Then
        MsgBox "Photo Column have now been deleted.", vbInformation, "Kutools for Excel"
    Else
        MsgBox "There are no Columns to delete as each one has more data (rows) than just a header.", vbExclamation, "Kutools for Excel"
    End If
    Application.ScreenUpdating = True
End Sub
