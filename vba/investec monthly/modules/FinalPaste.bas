Attribute VB_Name = "FinalPaste"
' FinalPaste
' Final step (Step04): copies calculated columns R:S from investec monthly.xlsm
' and pastes values into companies.xlsm starting at F2.
Option Explicit

Sub Step04FinalPaste()
    ' Copy calculated columns R:S from investec monthly and paste values into companies column F
    Windows("investec monthly.xlsm").Activate
    Range("R2:S7").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
