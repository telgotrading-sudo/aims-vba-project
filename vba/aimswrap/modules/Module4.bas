Attribute VB_Name = "Module4"
Sub Step02AIntegrity()
Attribute Step02AIntegrity.VB_ProcData.VB_Invoke_Func = " \n14"

'
    Sheets("aims").Select
    Application.CutCopyMode = False
    Range("D1502").Select
    ActiveCell.FormulaR1C1 = "=R[-1500]C"
    Range("F1502").Select
    ActiveCell.FormulaR1C1 = "=1*R[-1500]C"
    Range("D1502:F1502").Select
    Selection.Copy
    Range("D1502:D2817").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("F2:F1317").Select
    Selection.NumberFormat = "General"
    Range("F1502:F2817").Select
    Selection.Copy
End Sub
Sub Step02BIntegrityWrapToDel()

'
    Sheets("aimswrap").Select
    Application.CutCopyMode = False
    Range("A2:F380").Select
    Selection.NumberFormat = "General"
    Selection.Copy
    Range("A502").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("F502").Select
    ActiveCell.FormulaR1C1 = "=1*R[-500]C"
    Range("F502").Select
    Selection.Copy
    Range("F502:F880").Select
    ActiveSheet.Paste
    Selection.NumberFormat = "General"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub Step03ActiveSum()

Sheets("aimswrap").Select
Range("B2").Select
        
While ActiveCell.Value <> ""
    tmp01 = ActiveCell.Value
    Select Case ActiveCell().Offset(0, 3)
        Case "Stable SA"
            tmp01end = "a"
        Case "Global SA"
            tmp01end = "b"
        Case "Equities SA"
            tmp01end = "c"
        Case "Compulsory SA"
            tmp01end = "d"
        Case "Fairtree BCI Income Plus"
            tmp01end = "f"
        Case "Cash Movement"
            tmp01end = "k"
    End Select
    accno = tmp01 & tmp01end
    
    Sheets("aims").Select
    Range("B2").Select
    While ActiveCell.Value <> accno
        ActiveCell.Offset(1, 0).Select
    Wend
    activeval = ActiveCell.Offset(0, 4).Value
    While ActiveCell.Value <> ""
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value = accno Then
            activeval = activeval + ActiveCell.Offset(0, 4).Value
        End If
    Wend
    Sheets("aimswrap").Select
    ActiveCell.Offset(0, 4).Value = activeval
    ActiveCell.Offset(1, 0).Select
Wend
    
End Sub
Sub Step04FinalIntegrityCheckToDel()
'
'
'
    Sheets("aimswrap").Select

    Range("G502").Select
    ActiveCell().Offset(-1, 0).Activate
    ActiveCell.FormulaR1C1 = "%"
    
    ActiveCell().Offset(0, 1).Activate
    ActiveCell.FormulaR1C1 = "Concern"
    
    ActiveCell().Offset(1, -1).Activate
    ActiveCell.FormulaR1C1 = "=R[-500]C[-1]/RC[-1]-1"
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    
    ActiveCell().Offset(0, 1).Activate
    ActiveCell.FormulaR1C1 = "=IF(ABS(RC[-1])>0.1,""Check"","""")"
    
    Application.CutCopyMode = False
    Range("G502:H502").Select
    Selection.Copy
    Range("G502:G880").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
End Sub


