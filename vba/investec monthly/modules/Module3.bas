Attribute VB_Name = "Module3"
Sub Step04FinalPaste()
'
' MacroFS Macro
'

'
    Windows("investec monthly.xlsm").Activate
    Range("R2:S7").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
