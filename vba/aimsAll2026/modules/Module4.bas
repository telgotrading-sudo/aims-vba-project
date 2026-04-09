Attribute VB_Name = "Module4"
Sub Step05FinalPaste()
'
' MacroFS Macro
'

'
    Windows("aimsAll.xlsm").Activate
    Range("U2:V441").Select
    Selection.Copy
    Windows("aimswrap.xlsm").Activate
    Sheets("aimswrap").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
