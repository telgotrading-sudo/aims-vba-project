Attribute VB_Name = "ModuleImport1"
' ModuleImport
' Imports VBA components from the vba\<WorkbookName>\{modules|classes|forms}\
' directory tree back into .xlsm files in the excel\ folder.
' Prompts the user to choose a single workbook or import all at once.
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

    Dim fileList() As String
    Dim fileCount As Long
    Dim i As Long
    Dim choiceNum As Long
    Dim prompt As String
    Dim choice As String

    ' Root of the project — update this path if the repo moves
    projectRoot = "C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project"
    excelPath = projectRoot & "\excel\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(excelPath)

    ' First pass: collect all .xlsm filenames
    fileCount = 0
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
            fileCount = fileCount + 1
            ReDim Preserve fileList(1 To fileCount)
            fileList(fileCount) = file.Name
        End If
    Next file

    If fileCount = 0 Then
        MsgBox "No .xlsm files found in " & excelPath, vbExclamation
        Exit Sub
    End If

    ' Build the selection menu
    prompt = "Which workbook would you like to import into?" & vbCrLf & vbCrLf
    prompt = prompt & "  0 - All files" & vbCrLf
    For i = 1 To fileCount
        prompt = prompt & "  " & i & " - " & fileList(i) & vbCrLf
    Next i
    prompt = prompt & vbCrLf & "Enter a number (0 for all):"

    choice = InputBox(prompt, "Import Modules")
    If choice = "" Then Exit Sub  ' User cancelled

    If Not IsNumeric(choice) Then
        MsgBox "Invalid selection. Please enter a number.", vbExclamation
        Exit Sub
    End If

    choiceNum = CLng(choice)
    If choiceNum < 0 Or choiceNum > fileCount Then
        MsgBox "Please enter a number between 0 and " & fileCount & ".", vbExclamation
        Exit Sub
    End If

    ' Second pass: process selected file(s)
    For Each file In folder.Files

        If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then

            ' Skip files not matching the selection
            If choiceNum = 0 Or file.Name = fileList(choiceNum) Then

                ' Detect host workbook by name — use existing reference, skip Close
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

        End If

    Next file

    If choiceNum = 0 Then
        MsgBox "All imports complete!"
    Else
        MsgBox fileList(choiceNum) & " imported successfully!"
    End If
End Sub



