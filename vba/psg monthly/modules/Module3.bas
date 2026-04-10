Attribute VB_Name = "Module3"
Sub Step06PasteFinal()
Attribute Step06PasteFinal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MacroFS Macro
'

'
    Windows("psg monthly.xlsm").Activate
    Range("F2:G8").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
