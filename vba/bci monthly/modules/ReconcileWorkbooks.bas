Attribute VB_Name = "ReconcileWorkbooks"
' ReconcileWorkbooks
' Compares bci monthly.xlsm column B against the bci sheet in companies.xlsm column A
' row by row. Stops and selects the mismatching cells at the first difference found.
Option Explicit

Sub Step03MarkDiff()
    Dim wbMonthly As Workbook
    Dim wbCompanies As Workbook
    Dim wsMonthly As Worksheet
    Dim wsCompanies As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim bcitmp As String
    Dim comtmp As String

    On Error Resume Next
    Set wbMonthly = Workbooks("bci monthly.xlsm")
    Set wbCompanies = Workbooks("companies.xlsm")
    On Error GoTo 0

    If wbMonthly Is Nothing Or wbCompanies Is Nothing Then
        MsgBox "One or both workbooks not found.", vbExclamation
        Exit Sub
    End If

    Set wsMonthly = wbMonthly.ActiveSheet
    Set wsCompanies = wbCompanies.Sheets("bci")

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRow
        bcitmp = wsMonthly.Cells(i, "B").Value
        comtmp = wsCompanies.Cells(i, "A").Value

        ' Stop at the first row where values differ or companies data has ended
        If comtmp = "" Or comtmp <> bcitmp Then
            wbMonthly.Activate
            wsMonthly.Cells(i, "B").Select
            wbCompanies.Activate
            wsCompanies.Cells(i, "A").Select
            Exit Sub
        End If
    Next i

    MsgBox "No differences found.", vbInformation
End Sub
