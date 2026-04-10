Attribute VB_Name = "DataPreparation"
' DataPreparation
' Step01: reshapes the raw export on the active sheet by moving the header block,
' removing filler rows and columns, and writing clean column headings.
Option Explicit

Sub Step01PastePrep()
    ' Move the data block: cut A1:F1 and paste starting at C1, then drop columns A:B
    Range("A1:F1").Cut Destination:=Range("C1")
    Columns("A:B").Delete Shift:=xlToLeft

    ' Insert a blank row above the data and write clean column headers
    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "Fund"
    Range("B1").Value = "%"
    Range("C1").Value = "Date"
    Range("D1").Value = "Price"
    Range("E1").Value = "Units"
    Range("F1").Value = "Value"

    ' Autofit the date column and remove the two filler rows below the new header
    Columns("C:C").EntireColumn.AutoFit
    Rows("2:3").Delete Shift:=xlUp
End Sub
