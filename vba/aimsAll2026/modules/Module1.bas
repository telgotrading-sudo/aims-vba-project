Attribute VB_Name = "Module1"

Sub Step03markdiff()
Attribute Step03markdiff.VB_ProcData.VB_Invoke_Func = "N\n14"
    Dim wbAll As Workbook
    Dim wbWrap As Workbook
    Dim wsAll As Worksheet
    Dim wsWrap As Worksheet
    Dim lastRowAll As Long
    Dim lastRowWrap As Long
    Dim i As Long
    Dim aimsallcell As String
    Dim aimwrapcell As String
    Dim wrapEcell As String
    Dim allTcell As String
    Dim prevAimwrapcell As String
    
    ' Set references to workbooks and sheets
    On Error Resume Next
    Set wbAll = Workbooks("aimsAll.xlsm")
    Set wbWrap = Workbooks("aimswrap.xlsm")
    Set wsAll = wbAll.ActiveSheet
    Set wsWrap = wbWrap.Sheets("aimswrap")
    On Error GoTo 0
    
    ' Check if workbooks and sheet exist
    If wbAll Is Nothing Or wbWrap Is Nothing Or wsWrap Is Nothing Then
        MsgBox "One or both workbooks/sheets not found.", vbExclamation
        Exit Sub
    End If
    
    ' Find last rows in relevant columns
    lastRowAll = wsAll.Cells(wsAll.Rows.Count, "I").End(xlUp).Row
    lastRowWrap = wsWrap.Cells(wsWrap.Rows.Count, "B").End(xlUp).Row
    
    ' Start at row 2
    i = 2
    
    ' Loop through rows until a mismatch is found or data ends
    Do While i <= lastRowAll And i <= lastRowWrap
        ' Get values for comparison
        aimsallcell = wsAll.Cells(i, "I").Value
        aimwrapcell = Left(wsWrap.Cells(i, "B").Value, 10)
        wrapEcell = wsWrap.Cells(i, "E").Value
        allTcell = wsAll.Cells(i, "T").Value
        
        ' Get previous row's column B value in aimswrap (if available)
        If i > 2 Then
            prevAimwrapcell = Left(wsWrap.Cells(i - 1, "B").Value, 10)
        Else
            prevAimwrapcell = ""
        End If
        
        ' Check if column B is empty or values don't match
        If wsWrap.Cells(i, "B").Value = "" Or aimsallcell <> aimwrapcell Or wrapEcell <> allTcell Then
            ' Check if column R in aimsAll is "INVESTOR CHOICE"
            If wsAll.Cells(i, "R").Value = "INVESTOR CHOICE" Then
                ' Check if current or previous row in aimswrap column B matches aimsAll column I
                If aimsallcell = aimwrapcell Or aimsallcell = prevAimwrapcell Then
                    ' Determine which row to copy (current or previous)
                    Dim sourceRow As Long
                    If aimsallcell = aimwrapcell Then
                        sourceRow = i
                    Else
                        sourceRow = i - 1
                    End If
                    
                    wsWrap.Rows(sourceRow).Copy
                    wsWrap.Rows(i).Insert Shift:=xlDown
                                 
                    ' Copy column T from aimsAll to column E in the new row in aimswrap
                    wsWrap.Cells(i, "E").Value = wsAll.Cells(i, "T").Value
                    
                    ' Update last row in aimswrap due to insertion
                    lastRowWrap = wsWrap.Cells(wsWrap.Rows.Count, "B").End(xlUp).Row
                Else
                    ' No match found, select rows and exit
                    wsAll.Activate
                    wsAll.Cells(i, "I").Select
                    wbWrap.Activate
                    wsWrap.Cells(i, "B").Select
                    Exit Sub
                End If
            Else
                ' Non-INVESTOR CHOICE mismatch, select rows and exit
                wsAll.Activate
                wsAll.Cells(i, "I").Select
                wbWrap.Activate
                wsWrap.Cells(i, "B").Select
                Exit Sub
            End If
        End If
        
        i = i + 1
    Loop
    
    ' If loop ends due to reaching end of data, select the last checked rows
    If i <= lastRowAll Or i <= lastRowWrap Then
        wsAll.Activate
        wsAll.Cells(i, "I").Select
        wbWrap.Activate
        wsWrap.Cells(i, "B").Select
    End If
    
    ' If loop ends due to reaching end of data, show message
    If i > lastRowAll Or i > lastRowWrap Then
        MsgBox "No differences found.", vbInformation
    Else
        wsAll.Activate
        wsAll.Cells(i, "I").Select
        wbWrap.Activate
        wsWrap.Cells(i, "B").Select
    End If
End Sub
