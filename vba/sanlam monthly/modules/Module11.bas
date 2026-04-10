Attribute VB_Name = "Module11"
Sub Step01PastePrep()
Attribute Step01PastePrep.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("A1:F1").Select
    Selection.Cut
    Range("C1").Select
    ActiveSheet.Paste
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Fund"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "%"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Price"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Units"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Range("F2").Select
    Columns("C:C").EntireColumn.AutoFit
    
    Rows("2:3").Select
    Selection.Delete Shift:=xlUp
End Sub
