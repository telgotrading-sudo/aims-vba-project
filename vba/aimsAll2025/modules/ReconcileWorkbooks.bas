Attribute VB_Name = "ReconcileWorkbooks"
' ReconcileWorkbooks
' Compares aimsAll.xlsm and aimswrap.xlsm row-by-row to detect mismatches.
' For "INVESTOR CHOICE" rows, inserts a duplicate wrap row when needed.
' Stops at the first unresolvable mismatch and selects the differing cells.
Option Explicit

Sub Step03markdiff()
Attribute Step03markdiff.VB_ProcData.VB_Invoke_Func = "N\n14"
    Dim wbAll As Workbook
    Dim wbWrap As Workbook
    Dim wsAll As Worksheet
    Dim wsWrap As Worksheet
    Dim lastRowAll As Long
    Dim lastRowWrap As Long
    Dim i As Long
    Dim aimsAllPolicyCell As String
    Dim aimsWrapPolicyCell As String
    Dim wrapFundCell As String
    Dim allFundCell As String
    Dim prevAimsWrapPolicyCell As String
    Dim sourceRow As Long

    ' Set references to workbooks and sheets
    On Error Resume Next
    Set wbAll = Workbooks("aimsAll.xlsm")
    Set wbWrap = Workbooks("aimswrap.xlsm")
    Set wsAll = wbAll.ActiveSheet
    Set wsWrap = wbWrap.Sheets("aimswrap")
    On Error GoTo 0

    ' Check if workbooks and sheet exist
    If wbAll Is Nothing Or wbWrap Is Nothing Or wsWrap Is Nothing Then
        MsgBox "One or both workbooks/sheets not found.", vbExclamation
        Exit Sub
    End If

    ' Find last rows in relevant columns
    lastRowAll = wsAll.Cells(wsAll.Rows.Count, "I").End(xlUp).Row
    lastRowWrap = wsWrap.Cells(wsWrap.Rows.Count, "B").End(xlUp).Row

    ' Start at row 2 (row 1 is header)
    i = 2

    ' Loop through rows until a mismatch is found or data ends
    Do While i <= lastRowAll And i <= lastRowWrap
        ' Read values used for comparison
        aimsAllPolicyCell = wsAll.Cells(i, "I").Value
        aimsWrapPolicyCell = Left(wsWrap.Cells(i, "B").Value, 10)
        wrapFundCell = wsWrap.Cells(i, "E").Value
        allFundCell = wsAll.Cells(i, "T").Value

        ' Track the previous row's wrap policy number (for INVESTOR CHOICE lookahead)
        If i > 2 Then
            prevAimsWrapPolicyCell = Left(wsWrap.Cells(i - 1, "B").Value, 10)
        Else
            prevAimsWrapPolicyCell = ""
        End If

        ' Check if wrap row is missing or values don't match
        If wsWrap.Cells(i, "B").Value = "" Or aimsAllPolicyCell <> aimsWrapPolicyCell Or wrapFundCell <> allFundCell Then
            If wsAll.Cells(i, "R").Value = "INVESTOR CHOICE" Then
                ' INVESTOR CHOICE: try to resolve by inserting a duplicate wrap row
                If aimsAllPolicyCell = aimsWrapPolicyCell Or aimsAllPolicyCell = prevAimsWrapPolicyCell Then
                    ' Determine which existing wrap row to duplicate (current or previous)
                    If aimsAllPolicyCell = aimsWrapPolicyCell Then
                        sourceRow = i
                    Else
                        sourceRow = i - 1
                    End If

                    wsWrap.Rows(sourceRow).Copy
                    wsWrap.Rows(i).Insert Shift:=xlDown

                    ' Override fund name in the new row with aimsAll column T
                    wsWrap.Cells(i, "E").Value = wsAll.Cells(i, "T").Value

                    ' Recalculate last row after insertion
                    lastRowWrap = wsWrap.Cells(wsWrap.Rows.Count, "B").End(xlUp).Row
                Else
                    ' INVESTOR CHOICE but no matching policy found — highlight and stop
                    wsAll.Activate
                    wsAll.Cells(i, "I").Select
                    wbWrap.Activate
                    wsWrap.Cells(i, "B").Select
                    Exit Sub
                End If
            Else
                ' Non-INVESTOR CHOICE mismatch — highlight the differing cells and stop
                wsAll.Activate
                wsAll.Cells(i, "I").Select
                wbWrap.Activate
                wsWrap.Cells(i, "B").Select
                Exit Sub
            End If
        End If

        i = i + 1
    Loop

    ' Select the last checked row if data ended early in one workbook
    If i <= lastRowAll Or i <= lastRowWrap Then
        wsAll.Activate
        wsAll.Cells(i, "I").Select
        wbWrap.Activate
        wsWrap.Cells(i, "B").Select
    End If

    ' Notify if no differences were found through all rows
    If i > lastRowAll Or i > lastRowWrap Then
        MsgBox "No differences found.", vbInformation
    Else
        wsAll.Activate
        wsAll.Cells(i, "I").Select
        wbWrap.Activate
        wsWrap.Cells(i, "B").Select
    End If
End Sub
