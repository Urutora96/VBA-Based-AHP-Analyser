Attribute VB_Name = "Module2"
Sub ClearHighlight()
Attribute ClearHighlight.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.Offset(2, 0).Range("A1").Select
    Selection.Copy
    ActiveCell.Offset(-2, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
