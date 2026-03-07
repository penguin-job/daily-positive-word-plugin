Attribute VB_Name = "mod01_Main"
Option Explicit

Public Sub Run_ExecuteProcess()
    On Error GoTo EH
    
    Dim inputPath As String
    Dim outputFolder As String
    Dim outputPath As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1) Controlシートから取得
    inputPath = UI_GetInputPath()
    outputFolder = UI_GetOutputFolder()
    
    ' 2) 検証
    ValidateInputs inputPath, outputFolder
    
    ' 3) 出力パス生成
    outputPath = IO_BuildOutputPath(outputFolder, "processed_", "xlsx")
    
    ' 4) メイン処理（ここに中身を増やしていく）
    Biz_Process inputPath, outputPath
    
    ' 5) 結果を書き戻し（任意）
    UI_SetLastOutputPath outputPath
    
    MsgBox "処理完了：" & vbCrLf & outputPath, vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

EH:
    MsgBox "エラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub ValidateInputs(ByVal inputPath As String, ByVal outputFolder As String)
    If Len(Trim$(inputPath)) = 0 Then Err.Raise vbObjectError + 100, , "Inputファイルパスが未指定です。"
    If Len(Trim$(outputFolder)) = 0 Then Err.Raise vbObjectError + 101, , "Outputフォルダが未指定です。"
    
    If Not IO_FileExists(inputPath) Then Err.Raise vbObjectError + 102, , "Inputファイルが見つかりません：" & inputPath
    If Not IO_FolderExists(outputFolder) Then Err.Raise vbObjectError + 103, , "Outputフォルダが見つかりません：" & outputFolder
End Sub


