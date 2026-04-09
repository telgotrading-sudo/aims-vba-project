Attribute VB_Name = "Module2"
Sub Step04PasteFormulaDownToEditOrDel()
Attribute Step04PasteFormulaDownToEditOrDel.VB_ProcData.VB_Invoke_Func = " \n14"
'
' aimsccc Macro
'

'
    Windows("aimswrap.xlsm").Activate
    Sheets("aims").Select
    Range("F2:F1317").Select
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O2").Select
    Windows("aimswrap.xlsm").Activate
    Range("B2:B1317").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Q2").Select
    Windows("aimswrap.xlsm").Activate
    Range("H2:H1317").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F2").Select
    Windows("aimswrap.xlsm").Activate
    Range("E2:E1317").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("G2:M2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G3:G1317").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
End Sub
