Attribute VB_Name = "BalanceFilter"
' BalanceFilter
' Removes rows where the balance value in column I is below 5.
Option Explicit

Sub Step01BalDel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' Loop from bottom to top to safely delete rows without skipping
    For i = lastRow To 2 Step -1
        If IsNumeric(ws.Cells(i, "I").Value) And ws.Cells(i, "I").Value < 5 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
