Attribute VB_Name = "Module4"
Sub Step02MarkDiff()

Windows("investec monthly.xlsm").Activate
Range("C2").Select
mgtmp = ActiveCell().Value
Windows("companies.xlsm").Activate
Sheets("investec").Select
Range("A2").Select
comtmp = ActiveCell().Value

While ActiveCell().Value <> "" And Left(comtmp, 15) = Left(mgtmp, 15)

    Windows("investec monthly.xlsm").Activate
    ActiveCell().Offset(1, 0).Activate
    mgtmp = ActiveCell().Value
    Windows("companies.xlsm").Activate
    ActiveCell().Offset(1, 0).Activate
    comtmp = ActiveCell().Value

Wend

End Sub

