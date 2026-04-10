Attribute VB_Name = "NameFilter"
' NameFilter
' Removes rows for specific excluded client names from column A of the active sheet.
Option Explicit

Sub Step03NameDel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop from bottom to top to safely delete rows without skipping
    For i = lastRow To 2 Step -1
        cellValue = ws.Cells(i, "A").Value
        If cellValue = "Friederang, Corrinne" Or _
           cellValue = "Cavanagh, Gerald Ralph" Or _
           cellValue = "Maharaj, Saddarnun Madhkar" Or _
           cellValue = "Maggs, Roger Leonard Collis" Or _
           cellValue = "Pretorius, Magrieta Magdalena" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
