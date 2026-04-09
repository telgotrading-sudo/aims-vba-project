Attribute VB_Name = "Module3"
Sub compidfind()

  On Error GoTo EndMacro

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
    
Findstring = ActiveCell().Value
    
With Sheets("clientlist").Range("G:G")
    Set Rng = .Find(What:=Findstring, _
                    After:=.Cells(.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
    If Not Rng Is Nothing Then
        ActiveCell().Offset(0, -1).Value = Rng.Offset(0, -6).Value
        With ActiveCell().Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else
        With ActiveCell().Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End With

ActiveCell().Offset(1, 0).Activate

EndMacro:
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
  
End Sub


