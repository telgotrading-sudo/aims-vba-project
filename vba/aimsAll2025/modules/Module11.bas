Attribute VB_Name = "Module11"
Sub Step01bGeneralSort()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRowD As Long
    Dim lastRowI As Long
    Dim lastRowT As Long
    
    ' Set reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A (to determine the data range)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Find the last row in columns D, I, and T for sort keys
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    lastRowT = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    
    ' Clear existing sort fields
    ws.Sort.SortFields.Clear
    
    ' Add sort key for column D
    ws.Sort.SortFields.Add2 Key:=ws.Range("D2:D" & lastRowD), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Add sort key for column I
    ws.Sort.SortFields.Add2 Key:=ws.Range("I2:I" & lastRowI), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Add sort key for column T
    ws.Sort.SortFields.Add2 Key:=ws.Range("T2:T" & lastRowT), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Apply the sort to the entire data range (columns A to U)
    With ws.Sort
        .SetRange ws.Range("A1:U" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

