Attribute VB_Name = "PasteFormulas"
' PasteFormulas
' Step05: copies lookup data from companies.xlsm into psg monthly.xlsm,
' then fills formula columns F:K down from the row 2 template.
Option Explicit

Sub Step05CopyFormulasDown()
    ' Copy psgam sheet column H (prices) from companies.xlsm → psg monthly column N
    Windows("companies.xlsm").Activate
    Sheets("psgam").Select
    Range("H2:H8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("psg monthly.xlsm").Activate
    Range("N2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Copy psgam sheet column F from companies.xlsm → psg monthly column L
    Windows("companies.xlsm").Activate
    Range("F2:F8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("psg monthly.xlsm").Activate
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Copy psgam sheet column B (company names) from companies.xlsm → psg monthly column M
    Windows("companies.xlsm").Activate
    Range("B2:B8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("psg monthly.xlsm").Activate
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Fill formula columns F:K down from the row 2 template
    Application.CutCopyMode = False
    Range("F2:K2").Select
    Selection.Copy
    Range("F3:F8").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
