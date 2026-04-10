Attribute VB_Name = "Module2"
Sub Step05CopyFormulasDown()
Attribute Step05CopyFormulasDown.VB_ProcData.VB_Invoke_Func = " \n14"
'
' psgccc Macro
'

'
    Windows("companies.xlsm").Activate
    Sheets("psgam").Select
    Range("H2:H8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("psg monthly.xlsm").Activate
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("companies.xlsm").Activate
    Range("F2:F8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("psg monthly.xlsm").Activate
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("companies.xlsm").Activate
    Range("B2:B8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("psg monthly.xlsm").Activate
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("F2:K2").Select
    Selection.Copy
    Range("F3:F8").Select
    ActiveSheet.Paste
End Sub
