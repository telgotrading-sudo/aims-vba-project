Attribute VB_Name = "ReconcileWorkbooks"
' ReconcileWorkbooks
' Compares psg monthly.xlsm column C against the psgam sheet in companies.xlsm column B
' using the first 15 characters of each value. Stops and selects the mismatching cells
' at the first difference found.
Option Explicit

Sub Step04MarkDiff()
    Dim wbMonthly As Workbook
    Dim wbCompanies As Workbook
    Dim wsMonthly As Worksheet
    Dim wsCompanies As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim mgtmp As String
    Dim comtmp As String

    On Error Resume Next
    Set wbMonthly = Workbooks("psg monthly.xlsm")
    Set wbCompanies = Workbooks("companies.xlsm")
    On Error GoTo 0

    If wbMonthly Is Nothing Or wbCompanies Is Nothing Then
        MsgBox "One or both workbooks not found.", vbExclamation
        Exit Sub
    End If

    Set wsMonthly = wbMonthly.ActiveSheet
    Set wsCompanies = wbCompanies.Sheets("psgam")

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, "C").End(xlUp).Row

    For i = 2 To lastRow
        mgtmp = wsMonthly.Cells(i, "C").Value
        comtmp = wsCompanies.Cells(i, "B").Value

        ' Compare first 15 characters; stop at first mismatch or when companies data ends
        If comtmp = "" Or Left(comtmp, 15) <> Left(mgtmp, 15) Then
            wbMonthly.Activate
            wsMonthly.Cells(i, "C").Select
            wbCompanies.Activate
            wsCompanies.Cells(i, "B").Select
            Exit Sub
        End If
    Next i

    MsgBox "No differences found.", vbInformation
End Sub
