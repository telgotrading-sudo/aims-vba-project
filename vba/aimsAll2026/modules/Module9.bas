Attribute VB_Name = "Module9"
Sub Step02dDeleteLowBalanceRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column U
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    
    ' Loop through rows from bottom to top to safely delete rows
    For i = lastRow To 2 Step -1
        ' Get the value in column U
        cellValue = ws.Cells(i, "U").Value
        
        ' Check if the value is numeric and less than 100
        If IsNumeric(cellValue) And cellValue < 100 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
Sub Step02cDeleteNameRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Loop through rows from bottom to top to safely delete rows
    For i = lastRow To 2 Step -1
        ' Check if column E is "JANINE" and column F is "SCHWARTZ" or
        ' column E is "Wynand Petrus" and column F is "Steyn" (case-insensitive)
        If (UCase(ws.Cells(i, "E").Value) = "JANINE" And UCase(ws.Cells(i, "F").Value) = "SCHWARTZ") Or _
           (UCase(ws.Cells(i, "E").Value) = "WYNAND PETRUS" And UCase(ws.Cells(i, "F").Value) = "STEYN") Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

