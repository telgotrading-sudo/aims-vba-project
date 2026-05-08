Attribute VB_Name = "ReconcileWorkbooks"
' ReconcileWorkbooks
' Step02: compares MG monthly rows against the moriah global sheet in companies.xlsm.
' Handles known BCP cash account row differences before stopping at a mismatch.
Option Explicit

Private Const MG_WORKBOOK As String = "MG monthly.xlsm"
Private Const COMPANIES_WORKBOOK As String = "companies.xlsm"
Private Const COMPANIES_SHEET As String = "moriah global"
Private Const BCP_CASH_ACCOUNT As String = "BCP cash account (USD)"

Sub Step02Amarkdiff()
    Dim wbMonthly As Workbook
    Dim wbCompanies As Workbook
    Dim wsMonthly As Worksheet
    Dim wsCompanies As Worksheet
    Dim monthlyRow As Long
    Dim companiesRow As Long
    Dim mgCode As String
    Dim mgFund As String
    Dim companyCode As String
    Dim companyFund As String

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

    monthlyRow = 2
    companiesRow = 2

    Do While wsCompanies.Cells(companiesRow, "B").Value <> ""
        mgCode = wsMonthly.Cells(monthlyRow, "J").Value
        mgFund = wsMonthly.Cells(monthlyRow, "O").Value
        companyCode = wsCompanies.Cells(companiesRow, "B").Value
        companyFund = Left$(wsCompanies.Cells(companiesRow, "E").Value, 25)

        If companyCode = mgCode And InStr(1, LCase$(mgFund), LCase$(companyFund), vbTextCompare) <> 0 Then
            monthlyRow = monthlyRow + 1
            companiesRow = companiesRow + 1
        ElseIf companyCode = wsMonthly.Cells(monthlyRow - 1, "J").Value And companyFund = BCP_CASH_ACCOUNT Then
            wsCompanies.Rows(companiesRow).Delete Shift:=xlUp
            monthlyRow = monthlyRow - 1
            If monthlyRow < 2 Then monthlyRow = 2
        ElseIf mgCode = wsCompanies.Cells(companiesRow - 1, "B").Value And mgFund = BCP_CASH_ACCOUNT Then
            wsCompanies.Rows(companiesRow).Insert Shift:=xlDown
            CopyRowWithoutClipboard wsCompanies, companiesRow - 1, companiesRow
            wsCompanies.Cells(companiesRow, "E").Value = BCP_CASH_ACCOUNT
            monthlyRow = monthlyRow - 1
            If monthlyRow < 2 Then monthlyRow = 2
        Else
            wbMonthly.Activate
            wsMonthly.Cells(monthlyRow, "J").Select
            wbCompanies.Activate
            wsCompanies.Cells(companiesRow, "B").Select
            Exit Sub
        End If
    Loop

    MsgBox "No differences found.", vbInformation
End Sub

Private Sub CopyRowWithoutClipboard(ByVal ws As Worksheet, ByVal sourceRow As Long, ByVal targetRow As Long)
    Dim lastCol As Long

    lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, lastCol)).FormulaR1C1 = _
        ws.Range(ws.Cells(sourceRow, 1), ws.Cells(sourceRow, lastCol)).FormulaR1C1
    ws.Rows(targetRow).RowHeight = ws.Rows(sourceRow).RowHeight

    CopyRowFormattingWithoutClipboard ws, sourceRow, targetRow, lastCol
End Sub

Private Sub CopyRowFormattingWithoutClipboard(ByVal ws As Worksheet, ByVal sourceRow As Long, ByVal targetRow As Long, ByVal lastCol As Long)
    Dim colIndex As Long

    For colIndex = 1 To lastCol
        CopyCellFormatting ws.Cells(sourceRow, colIndex), ws.Cells(targetRow, colIndex)
    Next colIndex
End Sub

Private Sub CopyCellFormatting(ByVal sourceCell As Range, ByVal targetCell As Range)
    With targetCell
        .NumberFormat = sourceCell.NumberFormat
        .HorizontalAlignment = sourceCell.HorizontalAlignment
        .VerticalAlignment = sourceCell.VerticalAlignment
        .WrapText = sourceCell.WrapText
        .Orientation = sourceCell.Orientation
        .AddIndent = sourceCell.AddIndent
        .IndentLevel = sourceCell.IndentLevel
        .ShrinkToFit = sourceCell.ShrinkToFit
        .ReadingOrder = sourceCell.ReadingOrder
        .Font.Name = sourceCell.Font.Name
        .Font.Size = sourceCell.Font.Size
        .Font.Bold = sourceCell.Font.Bold
        .Font.Italic = sourceCell.Font.Italic
        .Font.Underline = sourceCell.Font.Underline
        .Font.Color = sourceCell.Font.Color
        .Interior.Pattern = sourceCell.Interior.Pattern
        .Interior.Color = sourceCell.Interior.Color
        .Interior.TintAndShade = sourceCell.Interior.TintAndShade
    End With

    CopyBorders sourceCell, targetCell
End Sub

Private Sub CopyBorders(ByVal sourceCell As Range, ByVal targetCell As Range)
    Dim borderIndex As Variant

    For Each borderIndex In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With targetCell.Borders(borderIndex)
            .LineStyle = sourceCell.Borders(borderIndex).LineStyle
            .Weight = sourceCell.Borders(borderIndex).Weight
            .Color = sourceCell.Borders(borderIndex).Color
        End With
    Next borderIndex
End Sub
