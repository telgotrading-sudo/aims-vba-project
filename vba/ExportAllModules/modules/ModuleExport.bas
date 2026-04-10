Attribute VB_Name = "ModuleExport"
' ModuleExport
' Exports the VBA components of .xlsm files in the excel\ folder into the
' vba\<WorkbookName>\{modules|classes|forms}\ directory tree.
' Prompts the user to choose a single workbook or export all at once.
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
    prompt = "Which workbook would you like to export?" & vbCrLf & vbCrLf
    prompt = prompt & "  0 - All files" & vbCrLf
    For i = 1 To fileCount
        prompt = prompt & "  " & i & " - " & fileList(i) & vbCrLf
    Next i
    prompt = prompt & vbCrLf & "Enter a number (0 for all):"

    choice = InputBox(prompt, "Export Modules")
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

                ' Don't close the host workbook — it is running this macro
                If Not isHostWorkbook Then
                    wb.Close SaveChanges:=False
                End If

            End If

        End If

    Next file

    If choiceNum = 0 Then
        MsgBox "All exports complete!"
    Else
        MsgBox fileList(choiceNum) & " exported successfully!"
    End If
End Sub


Private Sub CreateFolder(fso As Object, path As String)
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub


