Attribute VB_Name = "DelRow8"
Sub DeleteRow8()
Attribute DeleteRow8.VB_ProcData.VB_Invoke_Func = " \n14"
    Rows("8:8").Select
    Selection.Delete Shift:=xlUp
    Range("D8").Select
End Sub
