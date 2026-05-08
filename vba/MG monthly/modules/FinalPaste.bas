Attribute VB_Name = "FinalPaste"
' FinalPaste
' Step04: copies calculated MG columns AM:AN into companies.xlsm starting at F2.
Option Explicit

Private Const MG_WORKBOOK As String = "MG monthly.xlsm"
Private Const COMPANIES_WORKBOOK As String = "companies.xlsm"
Private Const COMPANIES_SHEET As String = "moriah global"

Sub Step04FinalCopy()
    Dim wbMonthly As Workbook
    Dim wbCompanies As Workbook
    Dim wsMonthly As Worksheet
    Dim wsCompanies As Worksheet
    Dim lastMonthlyRow As Long

    On Error Resume Next
    Set wbMonthly = Workbooks(MG_WORKBOOK)
    Set wbCompanies = Workbooks(COMPANIES_WORKBOOK)
    On Error GoTo 0

    If wbMonthly Is Nothing Or wbCompanies Is Nothing Then
        MsgBox "One or both workbooks not found.", vbExclamation
        Exit Sub
    End If

    Set wsMonthly = wbMonthly.ActiveSheet
    Set wsCompanies = wbCompanies.Worksheets(COMPANIES_SHEET)

    lastMonthlyRow = wsMonthly.Cells(wsMonthly.Rows.Count, "AM").End(xlUp).Row
    If lastMonthlyRow < 2 Then Exit Sub

    wsCompanies.Range("F2").Resize(lastMonthlyRow - 1, 2).Value = wsMonthly.Range("AM2:AN" & lastMonthlyRow).Value
End Sub
