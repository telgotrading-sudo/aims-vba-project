Attribute VB_Name = "Module2"
Sub Step01Achaims()

Sheets("aims").Select
Range("B2").Select
tmp01 = Left(ActiveCell().Value, 10)
Select Case Right(ActiveCell().Value, 1)
    Case "a"
        tmp01full = "Stable SA"
    Case "b"
        tmp01full = "Global SA"
    Case "c"
        tmp01full = "Equities SA"
    Case "d"
        tmp01full = "Compulsory SA"
    Case "f"
        tmp01full = "Fairtree BCI Income Plus"
    Case "k"
        tmp01full = "Cash Movement"
End Select

found = True

While ActiveCell().Value <> "" And found = True

    found = False
    Sheets("aimswrap").Select
    Range("B2").Select
    tmp02 = ActiveCell().Value
    tmp02full = ActiveCell().Offset(0, 3).Value
    
    Do While ActiveCell().Value <> ""
        If tmp01 = tmp02 And tmp01full = tmp02full Then
            found = True
            Exit Do
        End If
        
        ActiveCell().Offset(1, 0).Activate
        tmp02 = ActiveCell().Value
        tmp02full = ActiveCell().Offset(0, 3).Value
    Loop
    
    Sheets("aims").Select
    ActiveCell().Offset(1, 0).Activate
    tmp01 = Left(ActiveCell().Value, 10)
    Select Case Right(ActiveCell().Value, 1)
        Case "a"
            tmp01full = "Stable SA"
        Case "b"
            tmp01full = "Global SA"
        Case "c"
            tmp01full = "Equities SA"
        Case "d"
            tmp01full = "Compulsory SA"
        Case "f"
            tmp01full = "Fairtree BCI Income Plus"
        Case "k"
            tmp01full = "Cash Movement"
    End Select
    
Wend

End Sub
Sub Step01Bchaimswrap()

Sheets("aimswrap").Select
Range("B2").Select
tmp01 = ActiveCell().Value
Select Case ActiveCell().Offset(0, 3)
    Case "Stable SA"
        tmp01end = "a"
    Case "Global SA"
        tmp01end = "b"
    Case "Equities SA"
        tmp01end = "c"
    Case "Compulsory SA"
        tmp01end = "d"
    Case "Fairtree BCI Income Plus"
        tmp01end = "f"
    Case "Cash Movement"
        tmp01end = "k"
End Select
tmp01full = tmp01 & tmp01end

found = True

While ActiveCell().Value <> "" And found = True

    found = False
    Sheets("aims").Select
    Range("B2").Select
    tmp02 = ActiveCell().Value
    
    Do While ActiveCell().Value <> ""
        If tmp01full = tmp02 Then
            found = True
            Exit Do
        End If
        ActiveCell().Offset(1, 0).Activate
        tmp02 = ActiveCell().Value
    Loop
    
    Sheets("aimswrap").Select
    ActiveCell().Offset(1, 0).Activate
    tmp01 = ActiveCell().Value
    Select Case ActiveCell().Offset(0, 3)
        Case "Stable SA"
            tmp01end = "a"
        Case "Global SA"
            tmp01end = "b"
        Case "Equities SA"
            tmp01end = "c"
        Case "Compulsory SA"
            tmp01end = "d"
        Case "Fairtree BCI Income Plus"
            tmp01end = "f"
        Case "Cash Movement"
            tmp01end = "k"
    End Select
    tmp01full = tmp01 & tmp01end
    
Wend

End Sub

