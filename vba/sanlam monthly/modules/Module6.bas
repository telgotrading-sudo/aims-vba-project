Attribute VB_Name = "Module6"
Sub Step03MarkDiff()

Windows("sanlam monthly.xlsm").Activate
Range("k2").Select
mgtmp = ActiveCell().Value
Windows("companies.xlsm").Activate
Range("e2").Select
comtmp = ActiveCell().Value

While ActiveCell().Value <> "" And comtmp = mgtmp

    Windows("sanlam monthly.xlsm").Activate
    ActiveCell().Offset(1, 0).Activate
    mgtmp = ActiveCell().Value
    Windows("companies.xlsm").Activate
    ActiveCell().Offset(1, 0).Activate
    comtmp = ActiveCell().Value

Wend

Windows("sanlam monthly.xlsm").Activate

End Sub

