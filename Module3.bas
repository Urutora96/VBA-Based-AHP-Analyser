Attribute VB_Name = "Module3"
Option Explicit
Sub ClearTable()
Attribute ClearTable.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("B6:B8").Select
    Selection.ClearContents
    Range("E6:P8").Select
    Selection.ClearContents

    Range("E19:S22").Select
    Selection.ClearContents

    Range("B34:B39").Select
    Selection.ClearContents
    Range("E34:P39").Select
    Selection.ClearContents
End Sub
