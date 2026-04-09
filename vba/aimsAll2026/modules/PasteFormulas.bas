Attribute VB_Name = "PasteFormulas"
' PasteFormulas
' Copies data ranges from aimswrap.xlsm into aimsAll.xlsm (Step 04).
' Pastes values only (no formulas) for columns F, B, H and E of aimswrap
' into the corresponding columns of aimsAll, then fills down formula columns G:M.
' Note: Row range 2:1317 is fixed to match the expected data size for this workflow.
Option Explicit

Sub Step04PasteFormulaDownToEditOrDel()
Attribute Step04PasteFormulaDownToEditOrDel.VB_ProcData.VB_Invoke_Func = " \n14"

    ' Copy aimswrap column F (fund names) → aimsAll column N
    Windows("aimswrap.xlsm").Activate
    Sheets("aims").Select
    Range("F2:F1317").Select
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Copy aimswrap column B (policy numbers) → aimsAll column O
    Range("O2").Select
    Windows("aimswrap.xlsm").Activate
    Range("B2:B1317").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Copy aimswrap column H → aimsAll column Q
    Range("Q2").Select
    Windows("aimswrap.xlsm").Activate
    Range("H2:H1317").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Copy aimswrap column E (fund values) → aimsAll column F
    Range("F2").Select
    Windows("aimswrap.xlsm").Activate
    Range("E2:E1317").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("aimsAll.xlsm").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Fill formula columns G:M down from row 2 template to the full data range
    Range("G2:M2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G3:G1317").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub
