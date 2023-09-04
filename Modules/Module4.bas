Attribute VB_Name = "Module4"
Sub ResultDeclared()
    Dim rCell As Range
    Dim cRow As Long, LastRow As Long
    LastRow = Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
    With Worksheets("Sheet1").Range("A1", Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp))
        Do
            Set c = .Find(what:="*Result Declared on*", After:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns _
            , SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            If Not c Is Nothing Then
                cRow = c.Row
                c.EntireRow.Delete
            End If
         Loop While Not c Is Nothing And cRow < LastRow
    End With
	MsgBox "Result Declared Dates Removed "
End Sub

