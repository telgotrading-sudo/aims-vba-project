Attribute VB_Name = "PasteFormulas"
' PasteFormulas
' Step03: copies lookup data from companies.xlsm into investec monthly.xlsm,
' then fills formula columns R:V down from the row 2 template.
Option Explicit

Sub Step03CopyFormulasDown()
    ' Copy investec sheet column F (prices) from companies.xlsm → investec monthly column W
    Windows("companies.xlsm").Activate
    Sheets("investec").Select
    Range("F2:F7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("investec monthly.xlsm").Activate
    Range("W2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Copy investec sheet column A (company names) from companies.xlsm → investec monthly column X
    Windows("companies.xlsm").Activate
    Range("A2:A7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("investec monthly.xlsm").Activate
    Range("X2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Fill formula columns R:V down from the row 2 template
    Application.CutCopyMode = False
    Range("R2:V2").Select
    Selection.Copy
    Range("R3:R7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
