Attribute VB_Name = "DataPreparation"
' DataPreparation
' Step01: sorts the active sheet data by column C ascending,
' preparing rows in the correct order for the rest of the workflow.
Option Explicit

Sub Step01PastePrep()
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
