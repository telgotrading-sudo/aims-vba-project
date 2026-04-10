Attribute VB_Name = "FinalPaste"
' FinalPaste
' Final step (Step04): copies calculated columns H:I from sanlam monthly.xlsm
' and pastes values into companies.xlsm starting at F2.
Option Explicit

Sub Step04FinalPaste()
    ' Copy calculated columns H:I from sanlam monthly and paste values into companies column F
    Windows("sanlam monthly.xlsm").Activate
    Range("H2:I7").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
