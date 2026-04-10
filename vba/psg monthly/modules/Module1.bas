Attribute VB_Name = "Module1"
Sub Step02BalDel()
Attribute Step02BalDel.VB_ProcData.VB_Invoke_Func = " \n14"

Range("e2").Select

While ActiveCell().Value <> ""

If (ActiveCell().Value > 100) Then
    ActiveCell().Offset(1, 0).Activate
Else
    ActiveCell().EntireRow.Delete
End If

Wend
    
End Sub
