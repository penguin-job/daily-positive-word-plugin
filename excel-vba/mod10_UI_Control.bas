Attribute VB_Name = "mod10_UI_Control"
Option Explicit

Private Const CONTROL_SHEET As String = "Control"
Private Const CELL_INPUT As String = "B2"
Private Const CELL_OUTPUT As String = "B3"
Private Const CELL_LASTOUTPUT As String = "B4" ' 任意：最後の出力先表示

Public Function UI_GetInputPath() As String
    UI_GetInputPath = Sheets(CONTROL_SHEET).Range(CELL_INPUT).Value
End Function

Public Function UI_GetOutputFolder() As String
    UI_GetOutputFolder = Sheets(CONTROL_SHEET).Range(CELL_OUTPUT).Value
End Function

Public Sub UI_SetLastOutputPath(ByVal outputPath As String)
    Sheets(CONTROL_SHEET).Range(CELL_LASTOUTPUT).Value = outputPath
End Sub

'--- ボタン用：入力ファイル選択（例：CSV想定。必要なら変更）
Public Sub UI_SelectInputFile()
    Dim p As String
    p = IO_PickFile("入力ファイルを選択", _
                "CSV/TSVファイル", _
                "*.csv;*.tsv")

    If Len(p) > 0 Then Sheets(CONTROL_SHEET).Range(CELL_INPUT).Value = p
End Sub

'--- ボタン用：出力フォルダ選択
Public Sub UI_SelectOutputFolder()
    Dim p As String
    p = IO_PickFolder("出力フォルダを選択")
    If Len(p) > 0 Then Sheets(CONTROL_SHEET).Range(CELL_OUTPUT).Value = p
End Sub

