Attribute VB_Name = "Module1"
Sub Step01baldel()

Range("I2").Select

While ActiveCell().Value <> ""

If ActiveCell().Value < 5 Then
    ActiveCell().EntireRow.Delete
Else
    ActiveCell().Offset(1, 0).Activate
End If

Wend

End Sub
