Attribute VB_Name = "Module5"
Sub Step01PastePrep()
'
' pasteprep Macro
'

' Step 1 Autosize
    
    Columns("A:J").Select
    Columns("A:J").EntireColumn.AutoFit
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft

End Sub

