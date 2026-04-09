Attribute VB_Name = "Module10"
Sub ExportAllModules()

    Dim comp As Object
    Dim projectRoot As String
    Dim modulesPath As String
    Dim classesPath As String
    Dim formsPath As String

    ' Get project root (go up from /excel to root folder)
    projectRoot = "C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project"

    ' Build paths
    modulesPath = projectRoot & "\vba\modules\"
    classesPath = projectRoot & "\vba\classes\"
    formsPath = projectRoot & "\vba\forms\"

    ' Export components
    For Each comp In ThisWorkbook.VBProject.VBComponents

        ' Skip document modules (Sheet1, ThisWorkbook)
        If comp.Type = 100 Then GoTo NextComp

        Select Case comp.Type
            Case 1 ' Standard modules
                comp.Export modulesPath & comp.Name & ".bas"
                
            Case 2 ' Class modules
                comp.Export classesPath & comp.Name & ".cls"
                
            Case 3 ' Forms
                comp.Export formsPath & comp.Name & ".frm"
        End Select

NextComp:
    Next comp

    MsgBox "Export complete!"

End Sub
