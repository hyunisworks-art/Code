'このモジュールは、SharePoint上のExcelファイルやCSVファイルからデータを取得するための関数を提供します。
'J12_GetSharePoint.basは、w2p_get.basの59行目で呼び出されています。たとえば、J12_GetSharePoint.GetCsvDataのように利用されています。

Option Explicit

Function GetExcelData(ByVal path_excel As String, ByRef sheet_names() As Variant)
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheet_name As Variant
    
    Set GetExcelData = Nothing
    
    Set wb = Workbooks.Add
    For Each sheet_name In sheet_names
        On Error Resume Next
        With wb
            .Queries.Add Name:=sheet_name, Formula:= _
                "let" & vbCrLf & "ソース = Excel.Workbook(Web.Contents(""" & path_excel & """), null, true)," & vbCrLf & _
                sheet_name & "_Sheet = ソース{[Item=""" & sheet_name & """,Kind=""Sheet""]}[Data]," & vbCrLf & _
                "昇格されたヘッダー数 = Table.PromoteHeaders(" & sheet_name & "_Sheet, [PromoteAllScalars=true]) in 昇格されたヘッダー数"
            
            .Connections.Add2 "クエリ - " & sheet_name, "ブック内の '" & sheet_name & "' クエリへの接続です。", _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sheet_name & ";Extended Properties=" _
                , """" & sheet_name & """", 6, True, False
        End With
        
        If Err.Number = 1004 Then
            Application.DisplayAlerts = False
            wb.Close
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
            MsgBox "Sharepoint上のリスト参照に失敗しました1。" & vbCrLf & "ネットワーク環境や参照権限に問題がないか、ご確認ください。"
            Exit Function
        End If
        On Error GoTo 0
        
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Cells.NumberFormatLocal = "@"
        With ws.ListObjects.Add(SourceType:=4, Source:=wb.Connections("クエリ - " & sheet_name), Destination:=ws.Range("$A$1")).TableObject
            .RowNumbers = False
            .PreserveFormatting = True
            .RefreshStyle = 1
            .AdjustColumnWidth = True
            .ListObject.DisplayName = sheet_name
            .Refresh
        End With
        ws.Name = sheet_name
        
        Call QueryUnlist(wb)
        
    Next sheet_name
    
    Set GetExcelData = wb
End Function

Function GetCsvData(ByVal path_csv As String, ByVal enc As String)

    Dim sheet_name As String
    Dim wb As Workbook

    ' 既存ロジック：ファイル名からシート名を作成
    sheet_name = Right(path_csv, Len(path_csv) - InStrRev(path_csv, "\"))
    If sheet_name = path_csv Then
        MsgBox "同期したエクスプローラからファイルを選択してください。"
        Exit Function
    End If
    sheet_name = Left(sheet_name, InStr(sheet_name, ".") - 1)

    '（元の処理）エンコード判定（num_enc を作る）
    'enc = LCase(enc)
    'If enc = "utf-8" Or enc = "utf8" Then
    '    num_enc = 65001
    'ElseIf enc = "shiftjis" Or enc = "sjis" Or enc = "shift-jis" Then
    '    num_enc = 932
    'Else
    '    MsgBox "エンコード形式の指定に誤りがあります。管理者に問い合わせてください。"
    '    End
    'End If

    '（修正）PowerQueryを使わないため、ここではエンコードは後段のTextStreamで扱う
    enc = LCase(enc)
    If Not (enc = "utf-8" Or enc = "utf8" Or enc = "shiftjis" Or enc = "sjis" Or enc = "shift-jis") Then
        MsgBox "エンコード形式の指定に誤りがあります。管理者に問い合わせてください。"
        Exit Function
    End If

    Set GetCsvData = Nothing
    Set wb = Workbooks.Add

    '（元の処理）PowerQueryでCSV読み込み
    'On Error Resume Next
    'With wb
    '    .Queries.Add Name:=sheet_name, Formula:= _
    '        "let ソース = Csv.Document(File.Contents(""" & path_csv & """), [Delimiter="","", Encoding=" & num_enc & ", QuoteStyle=QuoteStyle.None])," & vbCrLf & _
    '        "昇格されたヘッダー数 = Table.PromoteHeaders(ソース, [PromoteAllScalars=true]) in 昇格されたヘッダー数"
    '
    '    With .Worksheets(1).ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sheet_name & ";Extended Properties=""""", Destination:=Range("$A$1")).QueryTable
    '        .CommandType = xlCmdSql
    '        .CommandText = Array("SELECT * FROM [" & sheet_name & "]")
    '        .ListObject.DisplayName = sheet_name
    '        .Refresh BackgroundQuery:=False
    '    End With
    'End With
    '
    'If Err.Number = 1004 Then
    '    Application.DisplayAlerts = False
    '    wb.Close
    '    Application.DisplayAlerts = True
    '    Application.ScreenUpdating = True
    '    MsgBox "Sharepoint上のリスト参照に失敗しました。2" & vbCrLf & "ネットワーク環境や参照権限に問題がないか、ご確認ください。"
    '    End
    'End If
    'On Error GoTo 0

    '（修正）FSO + TextStreamでCSV読込（ローカル/オンライン同期パスでも安定）
    Dim fso As Object, ts As Object
    Dim line As String
    Dim row As Long, col As Long
    Dim fields As Variant
    Dim maxCol As Long
    Dim colCount As Long

    On Error GoTo EH

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path_csv) = False Then
        Application.DisplayAlerts = False
        wb.Close
        Application.DisplayAlerts = True
        MsgBox "CSVファイルが見つかりません。" & vbCrLf & path_csv
        Exit Function
    End If

    ' TextStream: -1=Unicode, 0=ASCII, -2=SystemDefault
    ' UTF-8はOpenTextFileでは直接指定できないため ADODB.Stream を使うのが確実だが、
    ' 現場の安定性重視で「ADODB.StreamでUTF-8/Shift-JISを読み→行分割→貼り付け」にする
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = IIf(enc = "utf-8" Or enc = "utf8", "utf-8", "shift_jis") '（修正）指定エンコードで読む
    stm.Open
    stm.LoadFromFile path_csv
    Dim allText As String
    allText = stm.ReadText(-1)
    stm.Close

    ' 改行コード吸収（CRLF/LF混在対策）
    allText = Replace(allText, vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)

    row = 0
    maxCol = 0

    Dim lines As Variant
    lines = Split(allText, vbLf)

    For row = LBound(lines) To UBound(lines)
        line = lines(row)
        If Len(line) = 0 Then
            '末尾の空行などは無視
        Else
            'ダブルクォート内のカンマを正しく処理するため、w2p_get.splitCsv関数を使用
            fields = w2p_get.splitCsv(line)
            
            '列数の最大値を更新
            colCount = UBound(fields) - LBound(fields) + 1
            If colCount > maxCol Then maxCol = colCount

            'splitCsvの結果を直接シートに書き込み（1始まりの配列→1始まりの列）
            For col = LBound(fields) To UBound(fields)
                wb.Worksheets(1).Cells(row + 1, col).Value = fields(col)
            Next col
        End If
    Next row

    wb.Worksheets(1).Name = sheet_name

    '（元の処理）QueryUnlist はPowerQuery前提なので不要
    'Call QueryUnlist(wb)

    Set GetCsvData = wb
    Exit Function

EH:
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
    MsgBox "CSV読み込みに失敗しました。" & vbCrLf & _
           "Err=" & Err.Number & vbCrLf & Err.Description
    End

End Function

Sub QueryUnlist(wb)
    '接続を断つ
    For Each q In wb.Queries
        q.Delete
    Next q
    
    For Each ws In wb.Worksheets
        For Each qt In ws.QueryTables
            qt.Delete
        Next qt
        For Each obj In ws.ListObjects
            obj.Unlist
        Next obj
    Next ws
    
    For Each con In wb.Connections
        If con.Name = "ThisWorkbookDataModel" Then
        Else
            con.Delete
        End If
    Next con
    
    Application.CommandBars("Queries and Connections").Visible = True
End Sub

Sub SetFileName(ByVal my_file As String, ByRef rng_tmp As Range, ByVal ext_exp As String, ByVal ext As String)
'指定フォルダ
    'If Left(my_file, Len("http://")) = "http://" Or Left(my_file, Len("https://")) = "https://" Or my_file = "" Then
    MsgBox "ネットワーク上のファイルを参照する場合" & vbCrLf & "同期済みのローカルファイルを参照してください。"
    
    Set file_dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With file_dialog
        .Filters.Clear
        .Filters.Add ext_exp, "*." & ext, 1
        .AllowMultiSelect = False
        .Title = ""
        .InitialFileName = "C:\"
    End With
    
    If file_dialog.Show = False Then
        End
    Else
        my_file = file_dialog.SelectedItems(1)
    End If
    'End If
    rng_tmp.Value = my_file
End Sub

Function GetFileName(ByVal my_file As String, ByRef rng_tmp As Range)
    If Left(my_file, Len("http://")) = "http://" Or Left(my_file, Len("https://")) = "https://" Then
        GetFileName = rng_tmp.Value
    Else
        GetFileName = my_file
    End If
End Function

Sub SaveAsRetry(ByVal wb As Workbook, ByVal path_file As String)
    Dim flg_do As Boolean
    Dim s_time As Date
    Dim ret As VbMsgBoxResult
    
    Application.DisplayAlerts = False
    flg_do = False
    s_time = Now()
    Do While Dir(path_file, vbNormal) = ""
        flg_do = True
        On Error Resume Next
        Call wb.SaveAs(path_file)
        On Error GoTo 0
        If DateDiff("n", Now(), s_time) > 1 Then
            MsgBox "以下のファイルの保存に失敗しました" & vbCrLf & path_file
            End
        End If
    Loop
    
    If flg_do = False Then
        ret = MsgBox("保存しようとしているファイル" & vbCrLf & Dir(path_file, vbNormal) & vbCrLf & "が既にファイルが存在します。" & vbCrLf & "上書きしますか？", vbYesNo)
        If ret = vbNo Then
            MsgBox "処理を中断しました。"
            End
        Else
            Call wb.SaveAs(path_file)
        End If
    End If
    Application.DisplayAlerts = True
End Sub
