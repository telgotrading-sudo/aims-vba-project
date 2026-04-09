Attribute VB_Name = "FinalPaste"
' FinalPaste
' Final step (Step 05): copies columns U:V from aimsAll.xlsm
' and pastes values into aimswrap.xlsm (sheet "aimswrap") starting at F2.
Option Explicit

Sub Step05FinalPaste()

    ' Copy calculated columns U:V from aimsAll and paste values into aimswrap column F
    Windows("aimsAll.xlsm").Activate
    Range("U2:V461").Select
    Selection.Copy
    Windows("aimswrap.xlsm").Activate
    Sheets("aimswrap").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
