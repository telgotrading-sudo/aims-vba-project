Attribute VB_Name = "ModuleImport"
' ModuleImport
' Imports VBA components from the vba\<WorkbookName>\{modules|classes|forms}\
' directory tree back into each .xlsm in the excel\ folder.
' Run this after editing .bas/.cls/.frm files on disk to push changes into Excel.
' Note: this workbook is included in its own import run. It uses ThisWorkbook
' to avoid re-opening itself and to skip the Close (which would kill the macro).
Option Explicit

Sub ImportAllWorkbooks()
    Dim wb As Workbook
    Dim fso As Object
    Dim folder As Object
    Dim file As Object

    Dim projectRoot As String
    Dim excelPath As String
    Dim vbaPath As String
    Dim modulesPath As String
    Dim classesPath As String
    Dim formsPath As String
    Dim wbName As String
    Dim fileItem As Object
    Dim comp As Object
    Dim isHostWorkbook As Boolean

    ' Root of the project — update this path if the repo moves
    projectRoot = "C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project"
    excelPath = projectRoot & "\excel\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(excelPath)

    For Each file In folder.Files

        If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then

            ' Detect host workbook by name (more reliable than comparing full paths)
            isHostWorkbook = (LCase(file.Name) = LCase(ThisWorkbook.Name))

            If isHostWorkbook Then
                Set wb = ThisWorkbook
            Else
                Set wb = Workbooks.Open(file.path)
            End If

            wbName = Replace(file.Name, ".xlsm", "")

            vbaPath = projectRoot & "\vba\" & wbName & "\"
            modulesPath = vbaPath & "modules\"
            classesPath = vbaPath & "classes\"
            formsPath = vbaPath & "forms\"

            ' Remove existing standard modules, class modules and forms
            For Each comp In wb.VBProject.VBComponents
                If comp.Type = 1 Or comp.Type = 2 Or comp.Type = 3 Then
                    wb.VBProject.VBComponents.Remove comp
                End If
            Next comp

            ' Import modules, classes and forms from disk
            If fso.FolderExists(modulesPath) Then
                For Each fileItem In fso.GetFolder(modulesPath).Files
                    wb.VBProject.VBComponents.Import fileItem.path
                Next fileItem
            End If

            If fso.FolderExists(classesPath) Then
                For Each fileItem In fso.GetFolder(classesPath).Files
                    wb.VBProject.VBComponents.Import fileItem.path
                Next fileItem
            End If

            If fso.FolderExists(formsPath) Then
                For Each fileItem In fso.GetFolder(formsPath).Files
                    wb.VBProject.VBComponents.Import fileItem.path
                Next fileItem
            End If

            ' Don't close the host workbook — it is running this macro
            If Not isHostWorkbook Then
                wb.Close SaveChanges:=True
            End If

        End If

    Next file

    MsgBox "All imports complete!"
End Sub


