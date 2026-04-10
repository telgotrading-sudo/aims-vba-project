Attribute VB_Name = "Module2"
Sub Step02namedel()

Range("A2").Select

While ActiveCell().Value <> ""

If ActiveCell().Value = "3D TREE ANIMATION  VISUAL EFFECTS CC" Then
    ActiveCell().EntireRow.Delete
Else
    ActiveCell().Offset(1, 0).Activate
End If

Wend

End Sub
Sub Step04CopyFormulasDown()
Attribute Step04CopyFormulasDown.VB_ProcData.VB_Invoke_Func = " \n14"
'
' bciccc Macro
'

'
    Windows("companies.xlsm").Activate
    Sheets("bci").Select
    Range("F2:F7").Select
    Selection.Copy
    Windows("bci monthly.xlsm").Activate
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K2").Select
    Windows("companies.xlsm").Activate
    Range("A2:A7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("bci monthly.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("M2:Q2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M3:M7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
