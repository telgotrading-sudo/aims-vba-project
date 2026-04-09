Attribute VB_Name = "ReconcileAimsWrap"
Option Explicit

' Bidirectional reconciliation between the "aims" and "aimswrap" sheets.
'
' Step01Achaims  : checks that every entry in "aims" exists in "aimswrap".
' Step01Bchaimswrap : checks that every entry in "aimswrap" exists in "aims".
'
' Both subs use a fund-suffix convention to form full account identifiers:
'   a = Stable SA, b = Global SA, c = Equities SA, d = Compulsory SA,
'   f = Fairtree BCI Income Plus, k = Cash Movement

' ---------------------------------------------------------------------------
' Returns the full fund name for a given single-letter suffix.
' Used to map the last character of an aims account code to a display name.
' ---------------------------------------------------------------------------
Private Function FundSuffixToName(ByVal suffix As String) As String
    Select Case suffix
        Case "a": FundSuffixToName = "Stable SA"
        Case "b": FundSuffixToName = "Global SA"
        Case "c": FundSuffixToName = "Equities SA"
        Case "d": FundSuffixToName = "Compulsory SA"
        Case "f": FundSuffixToName = "Fairtree BCI Income Plus"
        Case "k": FundSuffixToName = "Cash Movement"
    End Select
End Function

' ---------------------------------------------------------------------------
' Returns the single-letter suffix for a given full fund name.
' Inverse of FundSuffixToName.
' ---------------------------------------------------------------------------
Private Function FundNameToSuffix(ByVal fundName As String) As String
    Select Case fundName
        Case "Stable SA":                FundNameToSuffix = "a"
        Case "Global SA":                FundNameToSuffix = "b"
        Case "Equities SA":              FundNameToSuffix = "c"
        Case "Compulsory SA":            FundNameToSuffix = "d"
        Case "Fairtree BCI Income Plus": FundNameToSuffix = "f"
        Case "Cash Movement":            FundNameToSuffix = "k"
    End Select
End Function

' ---------------------------------------------------------------------------
' Checks that every row in "aims" (starting B2) has a matching entry in
' "aimswrap". A match requires both the base account number (first 10 chars)
' and the full fund name (column B+3 in aimswrap) to agree.
' ---------------------------------------------------------------------------
Sub Step01Achaims()

    Dim tmp01 As String
    Dim tmp01full As String
    Dim tmp02 As String
    Dim tmp02full As String
    Dim found As Boolean

    Sheets("aims").Select
    Range("B2").Select

    tmp01 = Left(ActiveCell.Value, 10)
    tmp01full = FundSuffixToName(Right(ActiveCell.Value, 1))

    found = True

    While ActiveCell.Value <> "" And found = True

        found = False

        ' Search aimswrap for a matching account + fund name pair
        Sheets("aimswrap").Select
        Range("B2").Select
        tmp02 = ActiveCell.Value
        tmp02full = ActiveCell.Offset(0, 3).Value

        Do While ActiveCell.Value <> ""
            If tmp01 = tmp02 And tmp01full = tmp02full Then
                found = True
                Exit Do
            End If
            ActiveCell.Offset(1, 0).Activate
            tmp02 = ActiveCell.Value
            tmp02full = ActiveCell.Offset(0, 3).Value
        Loop

        ' Advance to the next aims row
        Sheets("aims").Select
        ActiveCell.Offset(1, 0).Activate
        tmp01 = Left(ActiveCell.Value, 10)
        tmp01full = FundSuffixToName(Right(ActiveCell.Value, 1))

    Wend

End Sub

' ---------------------------------------------------------------------------
' Checks that every row in "aimswrap" (starting B2) has a matching entry in
' "aims". Constructs the full aims account code by appending the fund suffix
' to the base account number, then searches column B of "aims" for it.
' ---------------------------------------------------------------------------
Sub Step01Bchaimswrap()

    Dim tmp01 As String
    Dim tmp01end As String
    Dim tmp01full As String
    Dim tmp02 As String
    Dim found As Boolean

    Sheets("aimswrap").Select
    Range("B2").Select

    tmp01 = ActiveCell.Value
    tmp01end = FundNameToSuffix(CStr(ActiveCell.Offset(0, 3)))
    tmp01full = tmp01 & tmp01end

    found = True

    While ActiveCell.Value <> "" And found = True

        found = False

        ' Search aims for the constructed full account code
        Sheets("aims").Select
        Range("B2").Select
        tmp02 = ActiveCell.Value

        Do While ActiveCell.Value <> ""
            If tmp01full = tmp02 Then
                found = True
                Exit Do
            End If
            ActiveCell.Offset(1, 0).Activate
            tmp02 = ActiveCell.Value
        Loop

        ' Advance to the next aimswrap row
        Sheets("aimswrap").Select
        ActiveCell.Offset(1, 0).Activate
        tmp01 = ActiveCell.Value
        tmp01end = FundNameToSuffix(CStr(ActiveCell.Offset(0, 3)))
        tmp01full = tmp01 & tmp01end

    Wend

End Sub
