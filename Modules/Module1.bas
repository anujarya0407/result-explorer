Attribute VB_Name = "Module1"
Sub RemoveAllGrades()

Range("A1:T13422").Replace what:="(O)", Replacement:=""
Range("A1:T13422").Replace what:="(A+)", Replacement:=""
Range("A1:T13422").Replace what:="(A)", Replacement:=""
Range("A1:T13422").Replace what:="(B+)", Replacement:=""
Range("A1:T13422").Replace what:="(B)", Replacement:=""
Range("A1:T13422").Replace what:="(C)", Replacement:=""
Range("A1:T13422").Replace what:="(D)", Replacement:=""
Range("A1:T13422").Replace what:="(P)", Replacement:=""
Range("A1:T13422").Replace what:="(F)", Replacement:=""

MsgBox "All Grades Removed"
End Sub

