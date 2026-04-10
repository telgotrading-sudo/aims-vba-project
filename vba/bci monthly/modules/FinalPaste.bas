Attribute VB_Name = "FinalPaste"
' FinalPaste
' Final step (Step05): copies calculated columns N:O from bci monthly.xlsm
' and pastes values into companies.xlsm (bci sheet) starting at F2.
Option Explicit

Sub Step05PasteFinal()
    ' Copy calculated columns N:O from bci monthly and paste values into companies column F
    Windows("bci monthly.xlsm").Activate
    Range("N2:O7").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
