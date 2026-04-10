Attribute VB_Name = "ModuleExport"
' ModuleExport
' Exports the VBA components of every .xlsm in the excel\ folder into the
' vba\<WorkbookName>\{modules|classes|forms}\ directory tree.
' Run this after making changes in Excel to push the latest code to disk.
Option Explicit

Sub ExportAllWorkbooks()
    Dim wb As Workbook
    Dim comp As Object
    Dim fso As Object
    Dim folder As Object
    Dim file As Object

    Dim projectRoot As String
    Dim excelPath As String
    Dim basePath As String
    Dim modulesPath As String
    Dim classesPath As String
    Dim formsPath As String
    Dim wbName As String

    ' Root of the project — update this path if the repo moves
    projectRoot = "C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project"

    excelPath = projectRoot & "\excel\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(excelPath)

    For Each file In folder.Files

        If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then

            ' Uncomment to skip the control workbooks themselves:
            'If file.Name = "ExportAllModules.xlsm" Or file.Name = "ImportAllModules.xlsm" Then GoTo NextFile

            Set wb = Workbooks.Open(file.path)

            wbName = Replace(file.Name, ".xlsm", "")

            ' Build output paths for this workbook
            basePath = projectRoot & "\vba\" & wbName & "\"
            modulesPath = basePath & "modules\"
            classesPath = basePath & "classes\"
            formsPath = basePath & "forms\"

            ' Ensure all output folders exist
            CreateFolder fso, projectRoot & "\vba\"
            CreateFolder fso, basePath
            CreateFolder fso, modulesPath
            CreateFolder fso, classesPath
            CreateFolder fso, formsPath

            ' Export each component to the appropriate subfolder
            For Each comp In wb.VBProject.VBComponents
                ' Skip document-level components (ThisWorkbook, Sheet modules)
                If comp.Type <> 100 Then
                    Select Case comp.Type
                        Case 1: comp.Export modulesPath & comp.Name & ".bas"
                        Case 2: comp.Export classesPath & comp.Name & ".cls"
                        Case 3: comp.Export formsPath & comp.Name & ".frm"
                    End Select
                End If
            Next comp

            wb.Close SaveChanges:=False

        End If

NextFile:
    Next file

    MsgBox "All exports complete!"
End Sub


Private Sub CreateFolder(fso As Object, path As String)
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub
