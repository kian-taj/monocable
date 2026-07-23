Attribute VB_Name = "Module2"
Option Explicit

Public Sub ExportEntireVBAProject()
    Dim component As Object
    Dim project As Object
    Dim exportRoot As String
    Dim exportFolder As String
    Dim extension As String
    Dim safeName As String

    If ThisWorkbook.path = vbNullString Then
        MsgBox "Save the workbook before exporting.", vbExclamation
        Exit Sub
    End If

    exportRoot = ThisWorkbook.path & Application.PathSeparator & _
                 "VBA_Export_" & Format(Now, "yyyymmdd_hhnnss")

    CreateFolderIfMissing exportRoot

    Set project = ThisWorkbook.VBProject

    For Each component In project.VBComponents
        extension = GetComponentExtension(component.Type)
        exportFolder = exportRoot & Application.PathSeparator & _
                       GetComponentFolder(component.Type)

        CreateFolderIfMissing exportFolder

        safeName = MakeSafeFileName(component.Name)

        component.Export exportFolder & Application.PathSeparator & _
                         safeName & extension
    Next component

    ExportProjectReferences project, exportRoot
    ExportWorkbookMetadata project, exportRoot

    MsgBox "VBA project exported to:" & vbCrLf & exportRoot, vbInformation
End Sub

Private Function GetComponentExtension(ByVal componentType As Long) As String
    Select Case componentType
        Case 1      ' vbext_ct_StdModule
            GetComponentExtension = ".bas"
        Case 2      ' vbext_ct_ClassModule
            GetComponentExtension = ".cls"
        Case 3      ' vbext_ct_MSForm
            GetComponentExtension = ".frm"
        Case 100    ' vbext_ct_Document
            GetComponentExtension = ".cls"
        Case Else
            GetComponentExtension = ".txt"
    End Select
End Function

Private Function GetComponentFolder(ByVal componentType As Long) As String
    Select Case componentType
        Case 1
            GetComponentFolder = "Modules"
        Case 2
            GetComponentFolder = "ClassModules"
        Case 3
            GetComponentFolder = "Forms"
        Case 100
            GetComponentFolder = "ExcelObjects"
        Case Else
            GetComponentFolder = "Other"
    End Select
End Function

Private Sub CreateFolderIfMissing(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Function MakeSafeFileName(ByVal value As String) As String
    Dim invalidCharacters As Variant
    Dim character As Variant

    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each character In invalidCharacters
        value = Replace$(value, character, "_")
    Next character

    MakeSafeFileName = value
End Function

Private Sub ExportProjectReferences(ByVal project As Object, _
                                    ByVal exportRoot As String)
    Dim reference As Object
    Dim fileNumber As Integer
    Dim outputFile As String

    outputFile = exportRoot & Application.PathSeparator & "references.txt"
    fileNumber = FreeFile

    Open outputFile For Output As #fileNumber

    For Each reference In project.References
        Print #fileNumber, "Name: " & reference.Name
        Print #fileNumber, "Description: " & reference.Description
        Print #fileNumber, "GUID: " & reference.GUID
        Print #fileNumber, "Version: " & reference.Major & "." & reference.Minor

        On Error Resume Next
        Print #fileNumber, "Path: " & reference.FullPath
        Print #fileNumber, "Missing: " & reference.IsBroken
        On Error GoTo 0

        Print #fileNumber, String$(60, "-")
    Next reference

    Close #fileNumber
End Sub

Private Sub ExportWorkbookMetadata(ByVal project As Object, _
                                   ByVal exportRoot As String)
    Dim component As Object
    Dim fileNumber As Integer
    Dim outputFile As String

    outputFile = exportRoot & Application.PathSeparator & "project_manifest.txt"
    fileNumber = FreeFile

    Open outputFile For Output As #fileNumber

    Print #fileNumber, "Workbook: " & ThisWorkbook.FullName
    Print #fileNumber, "VBA project: " & project.Name
    Print #fileNumber, "Exported: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, ""

    For Each component In project.VBComponents
        Print #fileNumber, component.Name & _
              " | Type=" & ComponentTypeName(component.Type) & _
              " | Lines=" & component.CodeModule.CountOfLines
    Next component

    Close #fileNumber
End Sub

Private Function ComponentTypeName(ByVal componentType As Long) As String
    Select Case componentType
        Case 1: ComponentTypeName = "Standard Module"
        Case 2: ComponentTypeName = "Class Module"
        Case 3: ComponentTypeName = "UserForm"
        Case 100: ComponentTypeName = "Excel Document Module"
        Case Else: ComponentTypeName = "Unknown"
    End Select
End Function

