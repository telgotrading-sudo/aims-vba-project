Attribute VB_Name = "Module7"
Sub Step01PastePrep()
Attribute Step01PastePrep.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With ActiveSheet
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("C1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange Range("A1").CurrentRegion
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub
