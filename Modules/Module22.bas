Attribute VB_Name = "Module22"
Sub ShiftCell()
Range("C1:J1").Select
    Selection.Delete shift:=xlUp
	
	MsgBox "Cells Shifted Up "
End Sub
