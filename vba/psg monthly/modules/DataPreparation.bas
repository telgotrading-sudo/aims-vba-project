Attribute VB_Name = "DataPreparation"
' DataPreparation
' Step01: autofits and removes unneeded columns from the active sheet,
' preparing the layout before the balance filter and data steps.
Option Explicit

Sub Step01PastePrep()
    ' Autofit columns A:J then remove the columns not needed for the workflow
    Columns("A:J").EntireColumn.AutoFit
    Columns("G:I").Delete Shift:=xlToLeft
    Columns("E:E").Delete Shift:=xlToLeft
    Columns("C:C").Delete Shift:=xlToLeft
End Sub
