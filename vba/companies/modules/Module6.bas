Attribute VB_Name = "Module6"
Sub CashAccount()
Attribute CashAccount.VB_ProcData.VB_Invoke_Func = "K\n14"
'
' CashAccount Macro
'
' Keyboard Shortcut: Ctrl+Shift+K
'
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(0, 4).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Cash account (USD)"
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub
