Attribute VB_Name = "DataPreparation"
' DataPreparation
' Three preparation steps run on the active sheet at the start of the workflow:
'   Step01a — cleans and standardises fund names into column T
'   Step02a — calculates per-policy totals from column N into column U
'   Step02b — copies rows that have a column U total to the next sheet (summary sheet)
Option Explicit

Sub Step01aNewCleanFundNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim iColumnValue As String
    Dim kColumnValue As String
    Dim prevIValue As String
    Dim nextIValue As String
    Dim prevRValue As String
    Dim nextRValue As String
    Dim cleanedValue As String
    Dim firstRowEmpty As Boolean
    Dim firstColEmpty As Boolean
    Dim j As Long

    ' Set reference to the active worksheet
    Set ws = ActiveSheet

    ' Check if the first row is empty (artifact from raw export)
    firstRowEmpty = True
    For j = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Not IsEmpty(ws.Cells(1, j)) Then
            firstRowEmpty = False
            Exit For
        End If
    Next j

    ' Check if the first column is empty (artifact from raw export)
    firstColEmpty = True
    For j = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Not IsEmpty(ws.Cells(j, 1)) Then
            firstColEmpty = False
            Exit For
        End If
    Next j

    ' Remove the first header row only if it is empty
    If firstRowEmpty Then
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
    End If

    ' Remove the leftmost column only if it is empty
    If firstColEmpty Then
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
    End If

    Range("A1").Select

    ' Find the last row in column R (product/fund type)
    lastRow = ws.Cells(ws.Rows.Count, "R").End(xlUp).Row

    ' Build a clean fund name in column T for every data row
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "R").Value
        iColumnValue = ws.Cells(i, "I").Value
        kColumnValue = ws.Cells(i, "K").Value

        If InStr(cellValue, "Kanaan") = 1 And InStr(cellValue, "Wrap") = Len(cellValue) - 3 Then
            ' Case 1: "Kanaan <FundName> Wrap" — strip the prefix and suffix
            cleanedValue = Trim(Mid(cellValue, 7, Len(cellValue) - 10))
            ws.Cells(i, "T").Value = cleanedValue

        ElseIf cellValue = "INVESTOR CHOICE" Then
            ' Case 2a: Investor Choice policy — use the fund name from column K
            ws.Cells(i, "T").Value = kColumnValue

        ElseIf cellValue = "Tax Application" Then
            ' Case 2b: Tax Application row — inherit fund name from the adjacent policy row
            If i > 1 Then
                prevIValue = ws.Cells(i - 1, "I").Value
                prevRValue = ws.Cells(i - 1, "R").Value
            Else
                prevIValue = ""
                prevRValue = ""
            End If

            If i < lastRow Then
                nextIValue = ws.Cells(i + 1, "I").Value
                nextRValue = ws.Cells(i + 1, "R").Value
            Else
                nextIValue = ""
                nextRValue = ""
            End If

            ' Match by policy number (column I) to previous or next row
            If iColumnValue = prevIValue And prevRValue <> "" Then
                If InStr(prevRValue, "Kanaan") = 1 And InStr(prevRValue, "Wrap") = Len(prevRValue) - 3 Then
                    cleanedValue = Trim(Mid(prevRValue, 7, Len(prevRValue) - 10))
                    ws.Cells(i, "T").Value = cleanedValue
                Else
                    ws.Cells(i, "T").Value = kColumnValue
                End If
            ElseIf iColumnValue = nextIValue And nextRValue <> "" Then
                If InStr(nextRValue, "Kanaan") = 1 And InStr(nextRValue, "Wrap") = Len(nextRValue) - 3 Then
                    cleanedValue = Trim(Mid(nextRValue, 7, Len(nextRValue) - 10))
                    ws.Cells(i, "T").Value = cleanedValue
                Else
                    ws.Cells(i, "T").Value = kColumnValue
                End If
            Else
                ' No adjacent policy match — fall back to column K
                ws.Cells(i, "T").Value = kColumnValue
            End If

        Else
            ' Case 2c: All other product types — use column K as-is
            ws.Cells(i, "T").Value = kColumnValue
        End If
    Next i
End Sub

Sub Step02aCalculatePolicyTotals()
    ' Aggregates column N values by policy (column I) + fund name (column T).
    ' Writes the group total into column U of the last row in each group.
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentPolicy As String
    Dim nextPolicy As String
    Dim currentName As String
    Dim nextName As String
    Dim total As Double

    ' Set reference to the active worksheet
    Set ws = ActiveSheet

    ' Find the last row in column I (policy numbers)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    total = 0

    For i = 2 To lastRow
        currentPolicy = ws.Cells(i, "I").Value
        currentName = ws.Cells(i, "T").Value

        If i < lastRow Then
            nextPolicy = ws.Cells(i + 1, "I").Value
            nextName = ws.Cells(i + 1, "T").Value
        Else
            nextPolicy = ""
            nextName = ""
        End If

        ' Accumulate the market value from column N
        If IsNumeric(ws.Cells(i, "N").Value) Then
            total = total + ws.Cells(i, "N").Value
        End If

        ' Write total when the policy/fund group ends
        If currentPolicy <> nextPolicy Or currentName <> nextName Or i = lastRow Then
            ws.Cells(i, "U").Value = total
            total = 0
        End If
    Next i
End Sub

Sub Step02bCopyRowsWithUValueToRightSheet()
    ' Copies the header row and all summary rows (those with a column U total)
    ' to the next sheet to the right, building a deduplicated summary sheet.
    Dim ws As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long

    ' Set reference to the active worksheet
    Set ws = ActiveSheet

    ' Copy header row to the target sheet first
    Rows("1:1").Select
    Selection.Copy

    On Error Resume Next
    Set wsTarget = ws.Next
    If wsTarget Is Nothing Then
        MsgBox "There is no sheet to the right of the current sheet.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    wsTarget.Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    ws.Select
    Range("A1").Select

    ' Find the last row in column T (fund name — present on all data rows)
    lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row

    ' Paste only the group-total rows (those with a value in column U)
    targetRow = 2
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, "U")) Then
            ws.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
            targetRow = targetRow + 1
        End If
    Next i
End Sub
