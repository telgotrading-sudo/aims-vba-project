Attribute VB_Name = "Module3"
Sub Step05PasteFinal()
Attribute Step05PasteFinal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MacroFS Macro
'

'
    Windows("bci monthly.xlsm").Activate
    Range("N2:O7").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
