Attribute VB_Name = "Module26"
Sub ShiftCell2()
Range("B2:I2").Select
    Selection.Delete shift:=xlUp
        
        MsgBox "Cells Shifted Up "
End Sub

