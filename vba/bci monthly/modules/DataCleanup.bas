Attribute VB_Name = "DataCleanup"
' DataCleanup
' Two cleanup steps for bci monthly.xlsm:
'   Step02NameDel  — removes the excluded company row from the active sheet
'   Step04CopyFormulasDown — copies lookup data from companies.xlsm and fills formula columns
Option Explicit

Sub Step02NameDel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop from bottom to top to safely delete rows without skipping
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "A").Value = "3D TREE ANIMATION  VISUAL EFFECTS CC" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub Step04CopyFormulasDown()
    ' Copy bci sheet column F (prices) from companies.xlsm → bci monthly column L
    Windows("companies.xlsm").Activate
    Sheets("bci").Select
    Range("F2:F7").Select
    Selection.Copy
    Windows("bci monthly.xlsm").Activate
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Copy bci sheet column A (company names) from companies.xlsm → bci monthly column K
    Windows("companies.xlsm").Activate
    Range("A2:A7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("bci monthly.xlsm").Activate
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Fill formula columns M:Q down from the row 2 template
    Application.CutCopyMode = False
    Range("M2:Q2").Select
    Selection.Copy
    Range("M3:M7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
