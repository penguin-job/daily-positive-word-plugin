Attribute VB_Name = "mod30_Biz_Process"
Option Explicit

' ==========================
' 入口：ワンクリック
'   入力：positive_words_100_jp_en.tsv（タブ区切り、ヘッダーあり：jp/en）
'   出力：quotes.tsv（タブ区切り、ヘッダーなし：jp/en）
' ==========================
Public Sub Publish_PositiveTSV_JpEn()
    Dim sourceUrl As String
    Dim sourceLocalPath As String

    ' ---- どちらか1つを使う ----
    ' A) URLから取り込む（WPメディアURL、GitHub raw、Dropbox直リンクなど）
    'sourceUrl = ""  ' 例: "https://yourdomain.com/wp-content/uploads/2026/02/positive_words_100_jp_en.tsv"

    ' B) ローカルTSVから取り込む（同じフォルダにある等）
    'sourceLocalPath = ThisWorkbook.path & "\positive_words_100_jp_en_with_header.tsv"
    
    sourceLocalPath = UI_GetInputPath()


    Dim ws As Worksheet
    Set ws = EnsureSheet("data")

    ' 1) 取り込み（TSV）
    If sourceUrl <> "" Then
        ImportTsvFromUrlToSheet sourceUrl, ws, True
    Else
        ImportTsvFromFileToSheet sourceLocalPath, ws, True
    End If

    ' 2) チェック＆整形（jp/en）
    NormalizeAndValidate_JpEn ws

    ' 3) 重複削除（jp列基準）
    RemoveDuplicatesByJp ws

    ' 4) ソート（任意：jp）
    SortData_JpEn ws

    ' 5) WPアップ用TSVを書き出し（UTF-8 BOM付き / ヘッダーなし）
    ExportCsvUtf8Bom_JpEn ws, ThisWorkbook.path & "\quotes.csv"


    MsgBox "完了：data を整形し、quotes.csv（jp/en, ヘッダーあり）を出力しました。"
End Sub


' ==========================
' 取り込み（URL）
' ==========================
Private Sub ImportTsvFromUrlToSheet(ByVal url As String, ByVal ws As Worksheet, ByVal clearSheet As Boolean)
    Dim txt As String
    txt = HttpGetTextUtf8(url)
    ImportTsvTextToSheet txt, ws, clearSheet
End Sub

' ==========================
' 取り込み（ローカルファイル）
' ==========================
Private Sub ImportTsvFromFileToSheet(ByVal filePath As String, ByVal ws As Worksheet, ByVal clearSheet As Boolean)
    Dim txt As String
    txt = ReadTextFileUtf8(filePath)
    ImportTsvTextToSheet txt, ws, clearSheet
End Sub

' 共通：TSVテキスト → シート（タブ区切り、最大2列）
Private Sub ImportTsvTextToSheet(ByVal tsvText As String, ByVal ws As Worksheet, ByVal clearSheet As Boolean)
    If clearSheet Then ws.Cells.Clear

    Dim normalized As String
    normalized = Replace(tsvText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim lines() As String
    lines = Split(normalized, vbLf)

    Dim r As Long: r = 1
    Dim i As Long

    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = lines(i)
        If Len(line) > 0 Then
            ' 末尾の空白だけ落とす（英語中の先頭スペース等を過剰に潰さない）
            line = RTrim$(line)
        End If

        If Trim$(line) <> "" Then
            Dim fields As Collection
            Set fields = ParseTsvLine(line)

            ws.Cells(r, 1).Value = IIf(fields.Count >= 1, fields(1), "")
            ws.Cells(r, 2).Value = IIf(fields.Count >= 2, fields(2), "")
            r = r + 1
        End If
    Next i

    ws.Columns("A:B").AutoFit
End Sub


' ==========================
' 整形＆チェック（jp/en）
' ==========================
Private Sub NormalizeAndValidate_JpEn(ByVal ws As Worksheet)
    ' 期待ヘッダー（ヘッダーあり運用）
    If LCase$(Trim$(ws.Cells(1, 1).Value)) <> "jp" Then ws.Cells(1, 1).Value = "jp"
    If LCase$(Trim$(ws.Cells(1, 2).Value)) <> "en" Then ws.Cells(1, 2).Value = "en"

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 700, , "データがありません。"

    Dim i As Long
    For i = 2 To lastRow
        ws.Cells(i, 1).Value = NormalizeText(ws.Cells(i, 1).Value)
        ws.Cells(i, 2).Value = NormalizeText(ws.Cells(i, 2).Value)
    Next i

    ' 空行（jpが空）を削除（下から）
    For i = lastRow To 2 Step -1
        If Trim$(CStr(ws.Cells(i, 1).Value)) = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    ws.Columns("A:B").AutoFit
End Sub

Private Function NormalizeText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(&H3000), " ") ' 全角スペース
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeText = s
End Function


' ==========================
' 重複削除（jp）
' ==========================
Private Sub RemoveDuplicatesByJp(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ws.Range("A1:B" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
End Sub


' ==========================
' ソート（任意：jp）
' ==========================
Private Sub SortData_JpEn(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    ws.Range("A1:B" & lastRow).Sort _
        Key1:=ws.Range("A2"), Order1:=xlAscending, _
        Header:=xlYes
End Sub


' ==========================
' TSV出力（UTF-8 BOM付き、ヘッダーあり：jp/en）
' ==========================
Private Sub ExportCsvUtf8Bom_JpEn(ByVal ws As Worksheet, ByVal outPath As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 710, , "出力するデータがありません。"

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    stm.Open

    ' ★ ヘッダー
    stm.WriteText "jp,en" & vbCrLf

    Dim r As Long
    For r = 2 To lastRow
        Dim jp As String, en As String
        jp = CStr(ws.Cells(r, 1).Value)
        en = CStr(ws.Cells(r, 2).Value)

        stm.WriteText CsvEscape(jp) & "," & CsvEscape(en) & vbCrLf
    Next r

    stm.SaveToFile outPath, 2 ' overwrite
    stm.Close
End Sub




' ==========================
' HTTP / ファイル読込（UTF-8）
' ==========================
Private Function HttpGetTextUtf8(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows; VBA)"
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 500, , "HTTP Error: " & http.Status & " " & url
    End If

    Dim bytes() As Byte
    bytes = http.ResponseBody

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    HttpGetTextUtf8 = stm.ReadText
    stm.Close
End Function

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUtf8 = stm.ReadText
    stm.Close
End Function


' ==========================
' TSVパーサ（タブ区切り）
' ==========================
Private Function ParseTsvLine(ByVal line As String) As Collection
    Dim result As New Collection
    Dim parts() As String
    Dim i As Long

    parts = Split(line, vbTab)

    For i = LBound(parts) To UBound(parts)
        result.Add CleanField(CStr(parts(i)))
    Next i

    Set ParseTsvLine = result
End Function

Private Function CleanField(ByVal s As String) As String
    ' 先頭のBOM（不可視文字）対策：UTF-8 BOM が混ざることがある
    If Len(s) > 0 Then
        ' U+FEFF (BOM) を除去
        If Left$(s, 1) = ChrW(&HFEFF) Then
            s = Mid$(s, 2)
        End If

        ' 文字化けで「i≫?」として入るケースも除去
        If Left$(s, 3) = "i≫?" Then
            s = Mid$(s, 4)
        End If
    End If

    s = Trim$(s)

    ' もし全体がダブルクォートで囲まれていれば外す
    If Len(s) >= 2 Then
        If Left$(s, 1) = """" And Right$(s, 1) = """" Then
            s = Mid$(s, 2, Len(s) - 2)
        End If
    End If

    CleanField = s
End Function

Private Function CsvEscape(ByVal s As String) As String
    Dim needQuote As Boolean
    needQuote = (InStr(s, ",") > 0) _
                Or (InStr(s, """") > 0) _
                Or (InStr(s, vbCr) > 0) _
                Or (InStr(s, vbLf) > 0)

    s = Replace(s, """", """""")

    If needQuote Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function


' ==========================
' シート作成/取得
' ==========================
Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set EnsureSheet = ws
End Function


