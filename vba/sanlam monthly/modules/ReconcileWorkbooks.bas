Attribute VB_Name = "ReconcileWorkbooks"
' ReconcileWorkbooks
' Compares sanlam monthly.xlsm column K against companies.xlsm column E row by row.
' Stops and selects the mismatching cells at the first difference found.
Option Explicit

Sub Step03MarkDiff()
    Dim wbMonthly As Workbook
    Dim wbCompanies As Workbook
    Dim wsMonthly As Worksheet
    Dim wsCompanies As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim mgtmp As String
    Dim comtmp As String

    On Error Resume Next
    Set wbMonthly = Workbooks("sanlam monthly.xlsm")
    Set wbCompanies = Workbooks("companies.xlsm")
    On Error GoTo 0

    If wbMonthly Is Nothing Or wbCompanies Is Nothing Then
        MsgBox "One or both workbooks not found.", vbExclamation
        Exit Sub
    End If

    Set wsMonthly = wbMonthly.ActiveSheet
    Set wsCompanies = wbCompanies.ActiveSheet

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, "K").End(xlUp).Row

    For i = 2 To lastRow
        mgtmp = wsMonthly.Cells(i, "K").Value
        comtmp = wsCompanies.Cells(i, "E").Value

        ' Stop at the first row where values differ or companies data has ended
        If comtmp = "" Or comtmp <> mgtmp Then
            wbMonthly.Activate
            wsMonthly.Cells(i, "K").Select
            wbCompanies.Activate
            wsCompanies.Cells(i, "E").Select
            Exit Sub
        End If
    Next i

    MsgBox "No differences found.", vbInformation
End Sub
