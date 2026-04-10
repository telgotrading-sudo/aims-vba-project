Attribute VB_Name = "Module5"
Sub Step04FinalPaste()

    Windows("sanlam monthly.xlsm").Activate
    Range("h2:i7").Select
    Selection.Copy
    Windows("companies.xlsm").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
