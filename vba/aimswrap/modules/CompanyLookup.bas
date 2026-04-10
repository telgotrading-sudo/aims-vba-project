Attribute VB_Name = "CompanyLookup"
Option Explicit

' Looks up a company ID (from the active cell) against column G of the "clientlist" sheet.
' - Match found    : writes the company name one column to the left and clears any red highlight.
' - No match found : highlights the active cell in red to flag the missing entry.
' After processing, advances the active cell down by one row.
' ...

Sub CompIdFind()

    Dim FindString As String
    Dim Rng As Range

    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    FindString = ActiveCell.Value

    With Sheets("clientlist").Range("G:G")
        Set Rng = .Find(What:=FindString, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
            ' Match found: populate company name and clear any red highlight
            ActiveCell.Offset(0, -1).Value = Rng.Offset(0, -6).Value
            With ActiveCell.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            ' No match: highlight active cell in red to flag the missing entry
            With ActiveCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    End With

    ActiveCell.Offset(1, 0).Activate

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
