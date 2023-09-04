Attribute VB_Name = "Module17"
Sub FindReplaceAllAsteriks()
'PURPOSE: Find & Replace text/values throughout a specific sheet


Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "~*"
rplc = ""

'Store a specfic sheet to a variable
  Set sht = Sheets("Sheet1")

'Perform the Find/Replace All
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
MsgBox "All Asteriks Are Now Removed"
End Sub

