Attribute VB_Name = "FinalPaste"
' FinalPaste
' Final step (Step06): copies calculated columns F:G from psg monthly.xlsm
' and pastes values into companies.xlsm starting at F2.
Option Explicit

Sub Step06PasteFinal()
    ' Copy calculated columns F:G from psg monthly and paste values into companies column F
    Windows("psg monthly.xlsm").Activate
    Range("F2:G8").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
