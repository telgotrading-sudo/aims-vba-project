Attribute VB_Name = "Module10"
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
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    ' Find the last row in column R
    lastRow = ws.Cells(ws.Rows.Count, "R").End(xlUp).Row
    
    ' Loop through each row in column R starting from R2
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "R").Value
        iColumnValue = ws.Cells(i, "I").Value
        kColumnValue = ws.Cells(i, "K").Value
        
        If InStr(cellValue, "Kanaan") = 1 And InStr(cellValue, "Wrap") = Len(cellValue) - 3 Then
            ' Case 1: Cell starts with "Kanaan" and ends with "Wrap"
            cleanedValue = Trim(Mid(cellValue, 7, Len(cellValue) - 10))
            ws.Cells(i, "T").Value = cleanedValue
        ElseIf cellValue = "INVESTOR CHOICE" Then
            ' Case 2a: Cell is "INVESTOR CHOICE"
            ws.Cells(i, "T").Value = kColumnValue
        ElseIf cellValue = "Tax Application" Then
            ' Case 2b: Cell is "Tax Application"
            ' Get values from column I and R for the previous and next rows
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
            
            ' Check if I column value matches previous or next row
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
                ' If no match is found, copy from column K
                ws.Cells(i, "T").Value = kColumnValue
            End If
        Else
            ' Case 2c: All other cases
            ws.Cells(i, "T").Value = kColumnValue
        End If
    Next i
End Sub
Sub Step02aCalculatePolicyTotals()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentPolicy As String
    Dim nextPolicy As String
    Dim currentName As String
    Dim nextName As String
    Dim total As Double
    Dim startRow As Long
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables
    total = 0
    startRow = 2
    
    ' Loop through each row in column I starting from I2
    For i = 2 To lastRow
        ' Get current and next row values
        currentPolicy = ws.Cells(i, "I").Value
        currentName = ws.Cells(i, "T").Value
        
        ' Get next row values, if available
        If i < lastRow Then
            nextPolicy = ws.Cells(i + 1, "I").Value
            nextName = ws.Cells(i + 1, "T").Value
        Else
            nextPolicy = ""
            nextName = ""
        End If
        
        ' Add current row's column N value to total
        If IsNumeric(ws.Cells(i, "N").Value) Then
            total = total + ws.Cells(i, "N").Value
        End If
        
        ' Check if the next row has a different policy number or name, or if it's the last row
        If currentPolicy <> nextPolicy Or currentName <> nextName Or i = lastRow Then
            ' Place the total in column U of the current row
            ws.Cells(i, "U").Value = total
            ' Reset total for the next group
            total = 0
        End If
    Next i
End Sub
Sub Step02bCopyRowsWithUValueToRightSheet()
    Dim ws As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    Rows("1:1").Select
    Selection.Copy
    
    ' Get the sheet to the right (next sheet)
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
    
    ' Find the last row in column T
    lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    
    ' Initialize target row for pasting (starting at row 2)
    targetRow = 2
    
    ' Loop through each row in column T starting from T2
    For i = 2 To lastRow
        ' Check if there is a value in column U
        If Not IsEmpty(ws.Cells(i, "U")) Then
            ' Copy the entire row to the target sheet
            ws.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
            ' Increment the target row for the next paste
            targetRow = targetRow + 1
        End If
    Next i
End Sub
