Attribute VB_Name = "SortWorksheet"
' SortWorksheet
' Sorts the active sheet data (columns A:U) by three keys in order:
'   1. Column D (date or client identifier — primary)
'   2. Column I (policy number — secondary)
'   3. Column T (fund name — tertiary)
' Run as Step01b immediately after Step01a (NewCleanFundNames).
Option Explicit

Sub Step01bGeneralSort()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRowD As Long
    Dim lastRowI As Long
    Dim lastRowT As Long

    ' Set reference to the active worksheet
    Set ws = ActiveSheet

    ' Find the last row in column A (full data extent)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Find last rows for each sort key column individually
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    lastRowT = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row

    ' Clear any previous sort definition
    ws.Sort.SortFields.Clear

    ' Add sort keys: column D (primary), I (secondary), T (tertiary)
    ws.Sort.SortFields.Add2 Key:=ws.Range("D2:D" & lastRowD), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    ws.Sort.SortFields.Add2 Key:=ws.Range("I2:I" & lastRowI), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    ws.Sort.SortFields.Add2 Key:=ws.Range("T2:T" & lastRowT), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    ' Apply the sort across the full data range A:U
    With ws.Sort
        .SetRange ws.Range("A1:U" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
