Attribute VB_Name = "Module6"
Sub Step03NameDel()

Range("A2").Select

While ActiveCell().Value <> ""

If ActiveCell().Value = "Friederang, Corrinne" Or ActiveCell().Value = "Cavanagh, Gerald Ralph" Or ActiveCell().Value = "Maharaj, Saddarnun Madhkar" Or ActiveCell().Value = "Maggs, Roger Leonard Collis" Or ActiveCell().Value = "Pretorius, Magrieta Magdalena" Then
    ActiveCell().EntireRow.Delete
Else
    ActiveCell().Offset(1, 0).Activate
End If

Wend

End Sub
