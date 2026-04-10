Attribute VB_Name = "PasteFormulas"
' PasteFormulas
' Step02: copies lookup data from companies.xlsm into sanlam monthly.xlsm,
' then fills formula columns G:M down from the row 2 template.
Option Explicit

Sub Step02CopyFormulasDown()
    ' Copy Sanlam sheet column F (prices) from companies.xlsm → sanlam monthly column N
    Windows("companies.xlsm").Activate
    Sheets("Sanlam").Select
    Range("F2:F7").Select
    Selection.Copy
    Windows("sanlam monthly.xlsm").Activate
    Range("N2").Select
    ActiveSheet.Paste

    ' Fill formula columns G:M down from the row 2 template
    Application.CutCopyMode = False
    Range("G2:M2").Select
    Selection.Copy
    Range("G3:G7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
