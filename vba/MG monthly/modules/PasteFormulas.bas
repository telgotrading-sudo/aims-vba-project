Attribute VB_Name = "PasteFormulas"
' PasteFormulas
' Step03: copies MG company lookup values from companies.xlsm and fills MG formulas.
Option Explicit

Private Const MG_WORKBOOK As String = "MG monthly.xlsm"
Private Const COMPANIES_WORKBOOK As String = "companies.xlsm"
Private Const COMPANIES_SHEET As String = "moriah global"

Sub Step03mgccc()
    Dim wbMonthly As Workbook
    Dim wbCompanies As Workbook
    Dim wsMonthly As Worksheet
    Dim wsCompanies As Worksheet
    Dim lastCompanyRow As Long

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

    lastCompanyRow = wsCompanies.Cells(wsCompanies.Rows.Count, "A").End(xlUp).Row
    If lastCompanyRow < 2 Then Exit Sub

    wsMonthly.Range("AR2").Resize(lastCompanyRow - 1, 1).Value = wsCompanies.Range("F2:F" & lastCompanyRow).Value
    wsMonthly.Range("AS2").Resize(lastCompanyRow - 1, 1).Value = wsCompanies.Range("A2:A" & lastCompanyRow).Value
    If lastCompanyRow > 2 Then
        wsMonthly.Range("AM2:AQ" & lastCompanyRow).FillDown
    End If
End Sub
