Attribute VB_Name = "shijisyo_create"

Sub create_shijisyo()
'==============================================
' 指示書CSV作成処理
'「W2Pデータ貼り付け」シートのテーブルを取得
Dim sheet_w2pdata As w2p_data
With ThisWorkbook.Worksheets(w2pdata_sheet)
    '最終列取得
    sheet_w2pdata.end_data_clm  = 39
    '最終行取得
    sheet_w2pdata.end_data_row = .Cells(.Rows.Count, 1).End(xlUp).Row

    '完全に空なら抜ける
    If sheet_w2pdata.end_data_row < 1 Or (sheet_w2pdata.end_data_row = 1 And .Cells(1, 1).Value = "") Then
        ReDim sheet_w2pdata.w2p_list(1 To 1, 1 To sheet_w2pdata.end_data_clm)
        sheet_w2pdata.w2p_list = Empty
        Exit Sub
    End If

    '配列の領域確保
    ReDim sheet_w2pdata.w2p_list(1 To sheet_w2pdata.end_data_row, 1 To sheet_w2pdata.end_data_clm)

    'シート→配列へ一括格納（高速）
    sheet_w2pdata.w2p_list = .Range(.Cells(1, 1), .Cells(sheet_w2pdata.end_data_row, sheet_w2pdata.end_data_clm)).Value
End With
        
'「W2Pデータ貼り付け」シートの列番号設定
with sheet_w2pdata
    .title_row = 1 'タイトル行
    .order_nom_clm = 2 '受注番号
    .haisou_order_name_clm = 8 '配送先名
    .item_code_clm = 20 '商品コード
    .item_name_clm = 21 '商品名
    .item_count_clm = 22 '数量
    .haisou_name_clm = 12 '配送先名
    .haisou_post_clm = 13 '配送先郵便番号
    .haisou_address1_clm = 14 '配送先住所1
    .haisou_address2_clm = 15 '配送先住所2
    .haisou_address3_clm = 16 '配送先住所3
    .haisou_tel_clm = 17 '配送先電話番号
    .haisousaki_tantou_clm = 18 '配送先担当者
    .nouki_clm = 32 '納期
    .sagyou_shiji_clm = 33 '作業指示
    .syukko_yotei_clm = 34 '出庫予定
    .end_data_clm = 0 '最終列（初期値0）
End With

'「ファイル名設定」シートの列番号設定と内容取得
Dim set_file_name_data As set_file_name_data
with set_file_name_data
    .start_row = 3
    .start_clm = 1
    .file_name_clm = 2
    .order_list_row = 1
    .shindou_list_row = 2
    .kyoten_list_row = 3
    .maru_list_row = 4
    .teikan_list_row = 7
    .end_row = 9
    .end_clm = 2
End With

'「ファイル名設定」シートの内容取得
With ThisWorkbook.Worksheets(set_file_name_sheet)
    set_file_name_data.file_name_list = .Range(.Cells(set_file_name_data.start_row, set_file_name_data.start_clm), .Cells(set_file_name_data.end_row, set_file_name_data.end_clm)).Value
End With

'W2Pデータ振り分け用変数初期化
Dim sep_w2p_data As sep_w2p_data
with sep_w2p_data
    .shindou_list_count = 1
    .maru_kyoten_list_count = 1
    .maru_list_count = 1
    .teikan_list_count = 1
    .shindou_list_title_row = 1
    .maru_kyoten_list_title_row = 1
    .maru_list_title_row = 1
    .teikan_list_title_row = 1
End With

'配列の領域確保
ReDim sep_w2p_data.shindou_list(1 To UBound(sheet_w2pdata.w2p_list, 1), 1 To UBound(sheet_w2pdata.w2p_list, 2))
Redim sep_w2p_data.maru_list(1 To UBound(sheet_w2pdata.w2p_list, 1), 1 To UBound(sheet_w2pdata.w2p_list, 2))
ReDim sep_w2p_data.maru_kyoten_list(1 To UBound(sheet_w2pdata.w2p_list, 1), 1 To UBound(sheet_w2pdata.w2p_list, 2))
ReDim sep_w2p_data.teikan_list(1 To UBound(sheet_w2pdata.w2p_list, 1), 1 To UBound(sheet_w2pdata.w2p_list, 2))

'タイトル行を各リストにコピー
Dim now_w2p_title_clm As Long
For now_w2p_title_clm = 1 To UBound(sheet_w2pdata.w2p_list, 2)
    sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_title_row, now_w2p_title_clm) = sheet_w2pdata.w2p_list(sheet_w2pdata.title_row, now_w2p_title_clm)
    sep_w2p_data.maru_list(sep_w2p_data.maru_list_title_row, now_w2p_title_clm) = sheet_w2pdata.w2p_list(sheet_w2pdata.title_row, now_w2p_title_clm)
    sep_w2p_data.maru_kyoten_list(sep_w2p_data.maru_kyoten_list_title_row, now_w2p_title_clm) = sheet_w2pdata.w2p_list(sheet_w2pdata.title_row, now_w2p_title_clm)
    sep_w2p_data.teikan_list(sep_w2p_data.teikan_list_title_row, now_w2p_title_clm) = sheet_w2pdata.w2p_list(sheet_w2pdata.title_row, now_w2p_title_clm)
Next now_w2p_title_clm

'パターン分けの色を指定
Dim patern_1_color As Long
Dim patern_2_color As Long
Dim patern_3_color As Long
Dim patern_4_color As Long

'色RGB値取得
patern_1_color = RGB(color1_R, color1_G, color1_B) '紫色
patern_2_color = RGB(color2_R, color2_G, color2_B) '緑色
patern_3_color = RGB(color3_R, color3_G, color3_B) '赤色
patern_4_color = RGB(color4_R, color4_G, color4_B) '黄色

'-----------------------------
'色に応じて配列に振り分ける（A列のセル色を確認して判定する）
Dim now_w2p_row As Long
Dim now_w2p_clm As Long
Dim cellColor As Long
'カウント初期化
sep_w2p_data.shindou_list_count = 1
sep_w2p_data.maru_kyoten_list_count = 1
sep_w2p_data.maru_list_count = 1
sep_w2p_data.teikan_list_count = 1

With ThisWorkbook.Worksheets(w2pdata_sheet)
    'データ行ループ
    For now_w2p_row = 2 To UBound(sheet_w2pdata.w2p_list, 1)
        'A列（1列目）のセル色を確認
        cellColor = .Cells(now_w2p_row, 1).Interior.Color
        If cellColor = patern_1_color Then
            'パターン1: 新藤Cに手配依頼するデータ（紫色）
            sep_w2p_data.shindou_list_count = sep_w2p_data.shindou_list_count + 1
            'データをコピー
            For now_w2p_clm = 1 To UBound(sheet_w2pdata.w2p_list, 2)
                sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_count, now_w2p_clm) = sheet_w2pdata.w2p_list(now_w2p_row, now_w2p_clm)
            Next
        ElseIf cellColor = patern_2_color Then
            'パターン2: マルテックスで商品をピックし、RLCが配送するデータ（緑色）
            sep_w2p_data.maru_kyoten_list_count = sep_w2p_data.maru_kyoten_list_count + 1
            For now_w2p_clm = 1 To UBound(sheet_w2pdata.w2p_list, 2)
                sep_w2p_data.maru_kyoten_list(sep_w2p_data.maru_kyoten_list_count, now_w2p_clm) = sheet_w2pdata.w2p_list(now_w2p_row, now_w2p_clm)
            Next
        ElseIf cellColor = patern_3_color Then
            'パターン3: マルテックスが配送まで手配するデータ（赤色）
            sep_w2p_data.maru_list_count = sep_w2p_data.maru_list_count + 1
            For now_w2p_clm = 1 To UBound(sheet_w2pdata.w2p_list, 2)
                sep_w2p_data.maru_list(sep_w2p_data.maru_list_count, now_w2p_clm) = sheet_w2pdata.w2p_list(now_w2p_row, now_w2p_clm)
            Next
        ElseIf cellColor = patern_4_color Then
            'パターン4: 定款（黄色）
            sep_w2p_data.teikan_list_count = sep_w2p_data.teikan_list_count + 1
            For now_w2p_clm = 1 To UBound(sheet_w2pdata.w2p_list, 2)
                sep_w2p_data.teikan_list(sep_w2p_data.teikan_list_count, now_w2p_clm) = sheet_w2pdata.w2p_list(now_w2p_row, now_w2p_clm)
            Next
        Else
            'どのパターンにも当てはまらない場合はパターン1（新藤）として扱う
            sep_w2p_data.shindou_list_count = sep_w2p_data.shindou_list_count + 1
            For now_w2p_clm = 1 To UBound(sheet_w2pdata.w2p_list, 2)
                sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_count, now_w2p_clm) = sheet_w2pdata.w2p_list(now_w2p_row, now_w2p_clm)
            Next
        End If
    Next
End With

'現在のブックのパス取得
Dim my_file As String
my_file = ThisWorkbook.Path & "\" & ThisWorkbook.Name

'現在のブックのディレクトリ取得
Dim my_dir As String
my_dir = Left(my_file, InStrRev(my_file, "\"))

'【受注データ csv】フォルダ作成
Dim cre_folder As String
cre_folder = my_dir & "\" & csv_folder_name
'既に存在している場合作成しない
If Dir(cre_folder, vbDirectory) = "" Then
    MkDir cre_folder
End If

'新藤様CSV：出力データが１件以上ある場合のみCSV書き出し
If sep_w2p_data.shindou_list_count >= 2 Then
    'ブックの新規作成:パターン１
    'ファイル名のYYYYMMDDを当日の日付に変換
    Dim file1_name As String
    file1_name = set_file_name_data.file_name_list(set_file_name_data.shindou_list_row, set_file_name_data.file_name_clm)
    file1_name = Replace(file1_name, "YYYYMMDD", Format(Now, "yyyymmdd"))
    file1_name = Replace(file1_name, "YYMMDD", Format(Now, "yymmdd"))
    Dim new_book1_path As String
    new_book1_path = cre_folder & "\" & file1_name
    ' CSV出力（共通関数を使用）
    Call ExportArrayToCsv(sep_w2p_data.shindou_list, new_book1_path, True)
End If

'マルテックスCSV：出力データが１件以上ある場合のみCSV書き出し
If sep_w2p_data.maru_list_count >= 2 Then
    'ブックの新規作成:パターン3
    'ファイル名のYYYYMMDDを当日の日付に変換
    Dim file3_name As String
    file3_name = set_file_name_data.file_name_list(set_file_name_data.maru_list_row, set_file_name_data.file_name_clm)
    file3_name = Replace(file3_name, "YYYYMMDD", Format(Now, "yyyymmdd"))
    file3_name = Replace(file3_name, "YYMMDD", Format(Now, "yymmdd"))
    Dim new_book3_path As String
    new_book3_path = cre_folder & "\" & file3_name
    ' CSV出力（共通関数を使用）
    Call ExportArrayToCsv(sep_w2p_data.maru_list, new_book3_path, True)
End If

'マルテックス拠点CSV：出力データが１件以上ある場合のみCSV書き出し    
If sep_w2p_data.maru_kyoten_list_count >= 2 Then
    'ブックの新規作成:パターン2
    'ファイル名のYYYYMMDDを当日の日付に変換
    Dim file2_name As String
    file2_name = set_file_name_data.file_name_list(set_file_name_data.kyoten_list_row, set_file_name_data.file_name_clm)
    file2_name = Replace(file2_name, "YYYYMMDD", Format(Now, "yyyymmdd"))
    file2_name = Replace(file2_name, "YYMMDD", Format(Now, "yymmdd"))
    Dim new_book2_path As String
    new_book2_path = cre_folder & "\" & file2_name
    ' CSV出力（共通関数を使用）
    Call ExportArrayToCsv(sep_w2p_data.maru_kyoten_list, new_book2_path, True)
End If
    
'定款CSVの出力：出力データが１件以上ある場合のみCSV書き出し
If sep_w2p_data.teikan_list_count >= 2 Then
    'ブックの新規作成:パターン4
    '定款ファイル名のYYYYMMDDを日付に変換
    Dim teikan_filename As String
    teikan_filename = set_file_name_data.file_name_list(set_file_name_data.teikan_list_row, set_file_name_data.file_name_clm)
    teikan_filename = Replace(teikan_filename, "YYYYMMDD", Format(Now, "yyyymmdd"))
    teikan_filename = Replace(teikan_filename, "YYMMDD", Format(Now, "yymmdd"))
    Dim teikan_folder_name As String
    teikan_folder_name = my_dir & "\" & csv_folder_name & "\" & teikan_folder
    '既に存在している場合、フォルダーは作成しない
    If Dir(teikan_folder_name, vbDirectory) = "" Then
        MkDir teikan_folder_name
    End If
    Dim new_book4_path As String
    new_book4_path = teikan_folder_name & "\" & teikan_filename
    ' CSV出力（共通関数を使用）
    Call ExportArrayToCsv(sep_w2p_data.teikan_list, new_book4_path, True)
End If

'作業指示フロー準備
Call prepInstrFlow(sep_w2p_data.shindou_list)

'作業指示書作成：出力データが１件以上ある場合のみ作成
If sep_w2p_data.shindou_list_count >= 2 Then
    '作業指示書作成
    call Create_Shijisho()
    'pdf出力
    Call print_shindo_pdf(set_file_name_data.file_name_list(1,2), my_dir, sep_w2p_data.shindou_list_count)
End If

'マルテックス用中間ファイルを出力
call middle_list_out(sep_w2p_data.maru_kyoten_list)
MsgBox "完了しました。"

End Sub

'==============================================
' CSV出力共通処理
' 配列データをCSVファイルに書き出す
'==============================================
Public Sub ExportArrayToCsv(ByRef dataArray As Variant, ByVal outputPath As String, _
                              Optional ByVal clearPriceColumns As Boolean = False, _
                              Optional ByVal charset As String = "Shift-JIS")
    
    Dim line As String
    Dim now_clm As Long
    Dim now_row As Long
    
    With CreateObject("ADODB.Stream")
        .charset = charset
        .Open
        
        ' ファイル名から価格列クリア対象か判定
        Dim outFileName As String
        Dim isTargetFile As Boolean
        outFileName = Mid(outputPath, InStrRev(outputPath, "\") + 1)
        If LCase(Right(outFileName, 4)) <> ".csv" Then outFileName = outFileName & ".csv"
        isTargetFile = (InStr(LCase(outFileName), "マルテックス") > 0)
        
        For now_row = 1 To UBound(dataArray, 1)
            If dataArray(now_row, 1) <> "" Then
                For now_clm = 1 To UBound(dataArray, 2)
                    ' 対象ファイルならデータ行の単価/小計を空にする
                    If isTargetFile And clearPriceColumns And now_row > 1 And (now_clm = 23 Or now_clm = 24) Then
                        dataArray(now_row, now_clm) = ""
                    End If
                    
                    ' 単価(23)/小計(24)を通貨表記に統一
                    If now_clm = 23 Or now_clm = 24 Then
                        If Trim(dataArray(now_row, now_clm)) <> "" Then
                            Dim v As String
                            v = dataArray(now_row, now_clm)
                            v = Replace(v, "\", "")
                            v = Replace(v, "\", "")
                            v = Replace(v, ",", "")
                            v = Replace(v, " ", "")
                            If IsNumeric(v) Then
                                dataArray(now_row, now_clm) = "\" & Format(CDbl(v), "#,##0.00")
                            End If
                        End If
                    End If
                    
                    ' CSVエスケープ処理
                    dataArray(now_row, now_clm) = Replace(dataArray(now_row, now_clm), Chr(34), Chr(34) & Chr(34))
                    dataArray(now_row, now_clm) = Chr(34) & dataArray(now_row, now_clm) & Chr(34)
                    If now_clm <> UBound(dataArray, 2) Then
                        dataArray(now_row, now_clm) = dataArray(now_row, now_clm) & ","
                    End If
                    line = line & dataArray(now_row, now_clm)
                Next
                .writetext line, 1
                line = ""
            End If
        Next
        
        .SaveToFile outputPath & ".csv", 2
        .Close
    End With
End Sub

'==============================================
' 作業指示フロー用データ準備
Sub prepInstrFlow(ByRef shindouList As Variant)
    '作業指示書作成リストシートの列番号設定
    Dim map As Variant
    map = Array(Array(2, 20), Array(8, 8), Array(20, 4), Array(21, 5), Array(22, 17), Array(13, 12), Array(14, 13), Array(15, 14), Array(16, 15), Array(17, 11), Array(18, 10), Array(32, 26), Array(34, 28),array(5,7),array(12,9),array(17,11))

    'シート保護を解除（画面更新なし）
    Worksheets(shijisyo_list_sheet).Unprotect

    'データ書き込み
    Dim i as Long
    dim k as Long
    Dim srcCol As Long
    Dim dstCol As Long
    '書き込み先シートへ
    With ThisWorkbook.Sheets(shijisyo_list_sheet)
        '既存データクリア（1行目は残す）
        .Rows("2:" & UBound(shindouList, 1)).ClearContents
        'データの書き込み
        For i = 2 To UBound(shindouList, 1)
            For k = LBound(map) To UBound(map)
                srcCol = map(k)(0)
                dstCol = map(k)(1)
                .Cells(i, dstCol).Value = CleanCsvVal(shindouList(i, srcCol))
            Next k
        Next i
    End With

    'シート保護を再設定
    Worksheets(shijisyo_list_sheet).Protect
End Sub



Sub print_shindo_pdf(ByVal pdf_file_name As String, ByVal my_dir As String, ByVal shijisyo_count As Long)
        '注文書作成
        With ThisWorkbook
            Dim f_name As String
            '注文書ファイル名設定
            f_name = pdf_file_name
            f_name = Replace(f_name, "YYYYMMDD", Format(Now, "yyyymmdd"))
            f_name = Replace(f_name, "YYMMDD", Format(Now, "yymmdd"))
            Dim order_count_up As Long
            '注文書のページ数計算（10件ずつのフォーマット）
            order_count_up = WorksheetFunction.Ceiling(shijisyo_count - 1, 10)
            Dim print_row As Long
            print_row = order_count_up * 3 + 5
            '注文書シートの印刷範囲設定
            Dim order_sh As Worksheet
            Set order_sh = .Worksheets(order_sheet)
            order_sh.PageSetup.PrintArea = order_sh.Range(order_sh.Cells(1, 1), order_sh.Cells(print_row, print_clm)).Address
            'PDF出力
            order_sh.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=my_dir & "\" & f_name
        End With
End Sub
'==============================================
