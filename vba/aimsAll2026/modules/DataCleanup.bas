Attribute VB_Name = "DataCleanup"
' DataCleanup
' Two cleanup passes run on the active sheet after initial data preparation:
'   Step02c — removes rows for excluded client names (Janine Schwartz, Wynand Petrus Steyn)
'   Step02d — removes rows where the policy total (column U) is below R100
Option Explicit

Sub Step02cDeleteNameRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set reference to the active worksheet
    Set ws = ActiveSheet

    ' Find the last row in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Loop from bottom to top to safely delete rows without skipping
    For i = lastRow To 2 Step -1
        ' Remove rows for excluded clients (case-insensitive match on first and last name)
        If (UCase(ws.Cells(i, "E").Value) = "JANINE" And UCase(ws.Cells(i, "F").Value) = "SCHWARTZ") Or _
           (UCase(ws.Cells(i, "E").Value) = "WYNAND PETRUS" And UCase(ws.Cells(i, "F").Value) = "STEYN") Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub Step02dDeleteLowBalanceRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant

    ' Set reference to the active worksheet
    Set ws = ActiveSheet

    ' Find the last row in column U (policy totals)
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row

    ' Loop from bottom to top to safely delete rows without skipping
    For i = lastRow To 2 Step -1
        cellValue = ws.Cells(i, "U").Value

        ' Remove rows where the policy total is numeric but below R100
        If IsNumeric(cellValue) And cellValue < 100 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
