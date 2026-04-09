Attribute VB_Name = "IntegrityWorkflow"
Option Explicit

' Four-step workflow that validates and summarises data between the
' "aims" and "aimswrap" sheets.
'
' Step 1  (ReconcileAimsWrap module) — bidirectional existence checks.
' Step 2A — extend "aims" formulas down to a staging area for comparison.
' Step 2B — copy "aimswrap" data to a staging area and convert to values.
' Step 3  — aggregate active totals from "aims" into "aimswrap" by account.
' Step 4  — add percentage-difference and flag formulas for final review.
'
' Subs whose names end in "ToDel" write temporary staging data that should
' be removed once validation is complete.

' ---------------------------------------------------------------------------
' Returns the single-letter fund suffix for a given full fund name.
' ---------------------------------------------------------------------------
Private Function FundNameToSuffix(ByVal fundName As String) As String
    Select Case fundName
        Case "Stable SA":                FundNameToSuffix = "a"
        Case "Global SA":                FundNameToSuffix = "b"
        Case "Equities SA":              FundNameToSuffix = "c"
        Case "Compulsory SA":            FundNameToSuffix = "d"
        Case "Fairtree BCI Income Plus": FundNameToSuffix = "f"
        Case "Cash Movement":            FundNameToSuffix = "k"
    End Select
End Function

' ---------------------------------------------------------------------------
' Extends a formula block in the "aims" sheet (rows 1502–2817) that mirrors
' rows 2–1317, allowing a numeric comparison of the two data sets.
' ---------------------------------------------------------------------------
Sub Step02AIntegrity()
Attribute Step02AIntegrity.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("aims").Select
    Application.CutCopyMode = False

    ' Seed the formula block at row 1502 and copy it down to row 2817
    Range("D1502").Select
    ActiveCell.FormulaR1C1 = "=R[-1500]C"
    Range("F1502").Select
    ActiveCell.FormulaR1C1 = "=1*R[-1500]C"
    Range("D1502:F1502").Select
    Selection.Copy
    Range("D1502:D2817").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    ' Convert the original value column to General format for comparison
    Range("F2:F1317").Select
    Selection.NumberFormat = "General"

    ' Copy the mirrored values for use in Step 2B
    Range("F1502:F2817").Select
    Selection.Copy

End Sub

' ---------------------------------------------------------------------------
' Copies the "aimswrap" data block (A2:F380) to a staging area starting at
' row 502, converts values to plain numbers, and pastes as values only.
' Note: writes temporary staging data — clear rows 502+ after validation.
' ---------------------------------------------------------------------------
Sub Step02BIntegrityWrapToDel()

    Sheets("aimswrap").Select
    Application.CutCopyMode = False

    ' Copy source data to staging area as values
    Range("A2:F380").Select
    Selection.NumberFormat = "General"
    Selection.Copy
    Range("A502").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    ' Force the F column to numeric by multiplying by 1, then paste as values
    Range("F502").Select
    ActiveCell.FormulaR1C1 = "=1*R[-500]C"
    Range("F502").Select
    Selection.Copy
    Range("F502:F880").Select
    ActiveSheet.Paste
    Selection.NumberFormat = "General"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

End Sub

' ---------------------------------------------------------------------------
' Aggregates active values from "aims" into "aimswrap" by account code.
' For each aimswrap row, constructs the full aims account code and sums all
' matching rows in "aims" column F (offset 4), writing the total back.
' ---------------------------------------------------------------------------
Sub Step03ActiveSum()

    Dim tmp01 As String
    Dim tmp01end As String
    Dim accno As String
    Dim activeval As Double

    Sheets("aimswrap").Select
    Range("B2").Select

    While ActiveCell.Value <> ""

        ' Build the full aims account code from base number + fund suffix
        tmp01 = ActiveCell.Value
        tmp01end = FundNameToSuffix(CStr(ActiveCell.Offset(0, 3)))
        accno = tmp01 & tmp01end

        ' Sum all matching rows in "aims"
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

        ' Write the aggregated total back to aimswrap
        Sheets("aimswrap").Select
        ActiveCell.Offset(0, 4).Value = activeval
        ActiveCell.Offset(1, 0).Select

    Wend

End Sub

' ---------------------------------------------------------------------------
' Adds a percentage-difference column (G) and a flag column (H) to the
' aimswrap staging area (rows 502–880) to surface discrepancies > 10%.
' Note: writes temporary staging data — clear rows 502+ after validation.
' ---------------------------------------------------------------------------
Sub Step04FinalIntegrityCheckToDel()

    Sheets("aimswrap").Select

    ' Write column headers in the row above the staging block
    Range("G502").Select
    ActiveCell.Offset(-1, 0).Activate
    ActiveCell.FormulaR1C1 = "%"
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.FormulaR1C1 = "Concern"

    ' Seed the percentage and flag formulas for row 502
    ActiveCell.Offset(1, -1).Activate
    ActiveCell.FormulaR1C1 = "=R[-500]C[-1]/RC[-1]-1"
    Selection.NumberFormat = "0.00%"
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.FormulaR1C1 = "=IF(ABS(RC[-1])>0.1,""Check"","""")"

    ' Copy both formula columns down through the full staging range
    Application.CutCopyMode = False
    Range("G502:H502").Select
    Selection.Copy
    Range("G502:G880").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub
