Attribute VB_Name = "mod20_IO_File"
Option Explicit

Public Function IO_PickFile(ByVal title As String, ByVal filterName As String, ByVal filterPattern As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add filterName, filterPattern
        
        If .Show = -1 Then
            IO_PickFile = .SelectedItems(1)
        Else
            IO_PickFile = vbNullString
        End If
    End With
End Function

Public Function IO_PickFolder(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .title = title
        If .Show = -1 Then
            IO_PickFolder = .SelectedItems(1)
        Else
            IO_PickFolder = vbNullString
        End If
    End With
End Function

Public Function IO_FileExists(ByVal path As String) As Boolean
    IO_FileExists = (Len(Dir$(path)) > 0)
End Function

Public Function IO_FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    IO_FolderExists = (Len(Dir$(folderPath, vbDirectory)) > 0)
    Exit Function
EH:
    IO_FolderExists = False
End Function

Public Function IO_CombinePath(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        IO_CombinePath = folderPath & fileName
    Else
        IO_CombinePath = folderPath & "\" & fileName
    End If
End Function

Public Function IO_BuildOutputPath(ByVal outputFolder As String, ByVal prefix As String, ByVal ext As String) As String
    Dim fileName As String
    fileName = prefix & Format(Now, "yyyymmdd_hhnnss") & "." & ext
    IO_BuildOutputPath = IO_CombinePath(outputFolder, fileName)
End Function

