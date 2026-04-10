Attribute VB_Name = "Module4"
Sub Step03markdiff()

Windows("bci monthly.xlsm").Activate
Range("B2").Select
bcitmp = ActiveCell().Value

Windows("companies.xlsm").Activate
Sheets("bci").Select
Range("A2").Select
comtmp = ActiveCell().Value



While ActiveCell().Value <> "" And comtmp = bcitmp

    Windows("bci monthly.xlsm").Activate
    ActiveCell().Offset(1, 0).Activate
    bcitmp = ActiveCell().Value
    Windows("companies.xlsm").Activate
    ActiveCell().Offset(1, 0).Activate
    comtmp = ActiveCell().Value
        
Wend

End Sub
