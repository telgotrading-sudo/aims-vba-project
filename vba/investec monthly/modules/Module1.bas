Attribute VB_Name = "Module1"
Sub Step03CopyFormulasDown()

'
' investecccc Macro
'

'
    Windows("companies.xlsm").Activate
    Sheets("investec").Select
    Range("F2:F7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("investec monthly.xlsm").Activate
    Range("W2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("companies.xlsm").Activate
    Range("A2:A7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("investec monthly.xlsm").Activate
    Range("X2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R2:V2").Select
    Selection.Copy
    Range("R3:R7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub
