Attribute VB_Name = "BalanceFilter"
' BalanceFilter
' Removes rows where the balance value in column E is 100 or below.
Option Explicit

Sub Step02BalDel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Loop from bottom to top to safely delete rows without skipping
    For i = lastRow To 2 Step -1
        If IsNumeric(ws.Cells(i, "E").Value) And ws.Cells(i, "E").Value <= 100 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
