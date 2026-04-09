Attribute VB_Name = "ModuleExport"
Option Explicit

' Exports all VBA components of this workbook to the file system.
' Standard modules (.bas), class modules (.cls), and forms (.frm) are
' written to the corresponding subfolders under the project root.
'
' Run this macro after making changes in the VBA editor to keep the
' repository files in sync with the workbook.

Sub ExportAllModules()

    Dim comp As Object
    Dim projectRoot As String
    Dim modulesPath As String
    Dim classesPath As String
    Dim formsPath As String

    ' Project root — adjust if the workbook is moved
    projectRoot = "C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project"

    modulesPath = projectRoot & "\vba\modules\"
    classesPath = projectRoot & "\vba\classes\"
    formsPath   = projectRoot & "\vba\forms\"

    For Each comp In ThisWorkbook.VBProject.VBComponents

        ' Skip document modules (Sheet objects, ThisWorkbook)
        If comp.Type = 100 Then GoTo NextComp

        Select Case comp.Type
            Case 1: comp.Export modulesPath & comp.Name & ".bas"  ' Standard module
            Case 2: comp.Export classesPath & comp.Name & ".cls"  ' Class module
            Case 3: comp.Export formsPath   & comp.Name & ".frm"  ' UserForm
        End Select

NextComp:
    Next comp

    MsgBox "Export complete!"

End Sub
