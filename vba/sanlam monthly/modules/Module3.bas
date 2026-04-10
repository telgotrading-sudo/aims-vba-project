Attribute VB_Name = "Module3"
Sub Step02CopyFormulasDown()

'
' sanlamccc Macro
'

'
    Windows("companies.xlsm").Activate
    Sheets("Sanlam").Select
    Range("f2:f7").Select
    Selection.Copy
    Windows("sanlam monthly.xlsm").Activate
    Range("n2:n2").Select
    ActiveSheet.Paste
    
    Range("g2:m2").Select
    Selection.Copy
    Range("g3:g7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub
