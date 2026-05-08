Attribute VB_Name = "DataPreparation"
' DataPreparation
' Step01: prepares the MG monthly export by sorting, removing low/invalid rows,
' and placing key fund rows in the expected order for reconciliation.
Option Explicit

Private Const MG_WORKBOOK As String = "MG monthly.xlsm"

Sub Step01PrepMG()
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ws.Columns("N:N").Delete Shift:=xlToLeft
    ws.Columns("O:P").EntireColumn.AutoFit

    SortByClientCode ws
    DeleteInternationalRows ws
    DeleteRowsBelowThreshold ws, "P", 0.2
    DeleteSmallNonCoreRows ws
    DeleteRowsBelowThreshold ws, "R", 0
    MarkAdditionalCashRows ws
    SortOrderedFundsNearTop ws
End Sub

Private Sub SortByClientCode(ByVal ws As Worksheet)
    With ws
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.Range("K1"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange ws.Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub

Private Sub DeleteInternationalRows(ByVal ws As Worksheet)
    Dim rowIndex As Long

    rowIndex = 350
    Do While ws.Cells(rowIndex, "A").Value <> ""
        If Left$(ws.Cells(rowIndex, "A").Value, 13) = "International" Then
            ws.Rows(rowIndex).Delete
        Else
            rowIndex = rowIndex + 1
        End If
    Loop
End Sub

Private Sub DeleteRowsBelowThreshold(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal threshold As Double)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim cellValue As Variant

    lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row

    For rowIndex = lastRow To 2 Step -1
        cellValue = ws.Cells(rowIndex, columnLetter).Value
        If IsNumeric(cellValue) Then
            If CDbl(cellValue) < threshold Then
                ws.Rows(rowIndex).Delete
            End If
        End If
    Next rowIndex
End Sub

Private Sub DeleteSmallNonCoreRows(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim balanceValue As Variant
    Dim fundName As String

    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row

    For rowIndex = lastRow To 2 Step -1
        balanceValue = ws.Cells(rowIndex, "P").Value
        fundName = ws.Cells(rowIndex, "O").Value

        If IsNumeric(balanceValue) Then
            If CDbl(balanceValue) < 2 And Not IsCoreFund(fundName) Then
                ws.Rows(rowIndex).Delete
            End If
        End If
    Next rowIndex
End Sub

Private Sub MarkAdditionalCashRows(ByVal ws As Worksheet)
    Dim rowIndex As Long
    Dim currentCode As String
    Dim previousCode As String
    Dim fundName As String
    Dim fundCount As Long

    rowIndex = 2
    Do While ws.Cells(rowIndex, "K").Value <> ""
        currentCode = ws.Cells(rowIndex, "K").Value
        If previousCode <> currentCode Then
            previousCode = currentCode
            fundCount = 0
        End If

        fundName = ws.Cells(rowIndex, "O").Value

        If Not IsOrderedFund(fundName) Then
            fundCount = fundCount + 1
            If fundCount > 1 Then
                With ws.Range(ws.Cells(rowIndex, "K"), ws.Cells(rowIndex, "L")).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End If

        rowIndex = rowIndex + 1
    Loop
End Sub

Private Sub SortOrderedFundsNearTop(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim priorityCol As Long
    Dim originalOrderCol As Long
    Dim rowIndex As Long

    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    priorityCol = lastCol + 1
    originalOrderCol = lastCol + 2

    ws.Cells(1, priorityCol).Value = "Fund sort priority"
    ws.Cells(1, originalOrderCol).Value = "Original row order"

    For rowIndex = 2 To lastRow
        ws.Cells(rowIndex, priorityCol).Value = FundSortPriority(CStr(ws.Cells(rowIndex, "O").Value))
        ws.Cells(rowIndex, originalOrderCol).Value = rowIndex
    Next rowIndex

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(1, "K"), ws.Cells(lastRow, "K")), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(1, priorityCol), ws.Cells(lastRow, priorityCol)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(1, originalOrderCol), ws.Cells(lastRow, originalOrderCol)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, originalOrderCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ws.Columns(priorityCol).Resize(, 2).Delete Shift:=xlToLeft
End Sub

Private Function FundSortPriority(ByVal fundName As String) As Long
    Select Case fundName
        Case "MCB cash account (USD)"
            FundSortPriority = 1
        Case "CIL Treasury account (USD)"
            FundSortPriority = 2
        Case "Moriah Global FoF (USD)"
            FundSortPriority = 3
        Case "Stable Offshore FoF (USD)"
            FundSortPriority = 4
        Case "Equity Offshore FoF (USD)"
            FundSortPriority = 5
        Case "Kanaan Income Offshore FoF (USD)"
            FundSortPriority = 6
        Case Else
            FundSortPriority = 99
    End Select
End Function

Private Function IsCoreFund(ByVal fundName As String) As Boolean
    IsCoreFund = fundName = "Stable Offshore FoF (USD)" _
        Or fundName = "Moriah Global FoF (USD)" _
        Or fundName = "Equity Offshore FoF (USD)" _
        Or fundName = "Kanaan Income Offshore FoF (USD)" _
        Or fundName = "CIL Treasury account (USD)" _
        Or fundName = "Currencies Offshore Fund (USD)"
End Function

Private Function IsOrderedFund(ByVal fundName As String) As Boolean
    IsOrderedFund = fundName = "Stable Offshore FoF (USD)" _
        Or fundName = "Moriah Global FoF (USD)" _
        Or fundName = "Equity Offshore FoF (USD)" _
        Or fundName = "CIL Treasury account (USD)" _
        Or fundName = "Kanaan Income Offshore FoF (USD)"
End Function
