Attribute VB_Name = "AccountSetup"
Option Explicit

' Inserts a copy of the current row directly below itself,
' then labels the new row's 5th column as "Cash account (USD)".
' Keyboard shortcut: Ctrl+Shift+K

Sub CashAccount()
Attribute CashAccount.VB_ProcData.VB_Invoke_Func = "K\n14"

    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    Selection.Insert Shift:=xlDown
    ActiveCell.Offset(0, 4).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Cash account (USD)"
    ActiveCell.Offset(1, 0).Range("A1").Select

End Sub
