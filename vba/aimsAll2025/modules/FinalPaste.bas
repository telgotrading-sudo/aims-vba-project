Attribute VB_Name = "FinalPaste"
' FinalPaste
' Final step (Step 05): copies columns U:V from aimsAll2025.xlsm
' and pastes values into aimswrap.xlsm (sheet "aimswrap") starting at F2.
Option Explicit

Sub Step05FinalPaste()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Copy calculated columns U:V from aimsAll2025 and paste values into aimswrap column F
    Windows("aimsAll2025.xlsm").Activate
    Set ws = ActiveSheet
    
    ' Find the last row with data in column U
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    
    ' Copy the range from row 2 to the last row with data
    Range("U2:V" & lastRow).Select
    Selection.Copy
    Windows("aimswrap.xlsm").Activate
    Sheets("aimswrap").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
