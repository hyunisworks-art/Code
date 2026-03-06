Attribute VB_Name = "w2p_get"
Option Explicit

'==========================================================
' メインプロシージャ: get_w2p
' 機能: CSVファイル選択→列数判定→W2P/Spinno処理振り分け
'       - 100列以上: W2P形式として process_w2p へ
'       - 15～50列: Spinno形式として process_spinno へ
'==========================================================
Sub get_w2p()
    
    '変数宣言
    Dim file_dialog As Office.FileDialog
    Dim my_dir As String
    Dim csv_path As String
    Dim csv_wb As Workbook
    Dim csv_data As csv_data
    Dim csvColCount As Long
    Dim lastRow As Long
    Dim lastClm As Long
    
    '確認メッセージ
    If MsgBox("既存データを削除して、データを取り込みます。よろしいですか？", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
        
    'ファイル選択ダイアログを表示
    my_dir = ThisWorkbook.path & "\" & ThisWorkbook.Name
    
    'ディレクトリパスの取得
    my_dir = Left(my_dir, InStrRev(my_dir, "\"))
    If Right(my_dir, 1) <> "\" Then
        my_dir = my_dir & "\"
    End If
    
    Set file_dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With file_dialog
        .Filters.Clear
        .Filters.Add "CSV", "*.csv", 1
        .FilterIndex = 1
        .AllowMultiSelect = False
        .Title = "取り込むCSVファイルを選択してください。"
        .InitialFileName = my_dir
    End With
        
    If file_dialog.Show = False Then
        End
    Else
        csv_path = file_dialog.SelectedItems(1)
    End If
    
    '取り込んだCSVファイルのパスを保存
    With ThisWorkbook.Worksheets(file_name_save_sheet)
        .Cells(1, 1).Value = csv_path
    End With
    
    'CSVデータをUTF-8で読み込む
    Set csv_wb = J12_GetSharePoint.GetCsvData(csv_path, "utf8")
    
    With csv_wb.Worksheets(1)
        lastRow = .Cells(1, 1).SpecialCells(xlLastCell).row
        lastClm = .Cells(1, 1).SpecialCells(xlLastCell).Column
        csv_data.csv_list = .Range(.Cells(1, 1), .Cells(lastRow, lastClm)).Value
    End With
        
    On Error Resume Next
    Application.DisplayAlerts = False
    csv_wb.Close
    Application.DisplayAlerts = True
    On Error GoTo 0
    '--- 列数判定して適切な処理へ振り分け ---
    csvColCount = UBound(csv_data.csv_list, 2)
    
    If csvColCount >= 100 Then
        'W2P形式（143列前後）→ process_w2p へ
        Call process_w2p(csv_data)
        
    ElseIf csvColCount >= 15 And csvColCount <= 50 Then
        'Spinno形式（20列前後）→ process_spinno へ
        Call process_spinno(csv_data)
        
    Else
        MsgBox "想定外のCSV列数(" & csvColCount & ")です。処理を中止します。" & vbCrLf & _
                "W2P形式（100列以上）またはSpinno形式（15～50列）のCSVを選択してください。"
        Exit Sub
    End If
    
End Sub

'==========================================================
' W2P処理プロシージャ（143列CSV用）
' 機能: W2P形式のCSVデータ（143列）を読み込み、
'       商品コード別に色分けし、納期・出荷予定日を計算
'==========================================================
Sub process_w2p(csv_data As csv_data)
    
    '変数宣言
    Dim patern_1_color As Long, patern_2_color As Long
    Dim patern_3_color As Long, patern_4_color As Long
    Dim w2p_data As w2p_data
    Dim t_code_data As teikan_code_data
    Dim haisousaki_data As haisousaki_data
    Dim set_file_name_data As set_file_name_data
    Dim hed_syukka() As Variant, hed_nyuuko() As Variant
    Dim clm_link_data As clm_link_data
    Dim end_row As Long, end_clm As Long
    Dim now_row As Long, now_clm As Long
    Dim lastUsedRow As Long
    Dim rMax As Long, cMax As Long
    Dim body() As Variant
    Dim r As Long, c As Long
    Dim now_order_clm As Long, now_address_clm As Long, now_w2p_clm As Long
    Dim match_flg As Boolean
    Dim match_key() As Variant
    Dim match_title_row As Long, match_address_row As Long, match_w2p_row As Long
    Dim circle_count As Long, now_key_clm As Long, now_link_clm As Long, w2p_clm As Long
    Dim RET_DAY As RET_DAY, send_time As Long, order_date As Date
    Dim to_send_days As Long, retern_day As RET_DAY, teikan_flg As Boolean
    Dim cs_data As color_send_days_data, ws As Worksheet
    Dim user_set_flg As Boolean, now_syouhin_code As Long, key As String
    Dim now_cs_data_row As Long, now_haisou_row As Long, match_key_clm As Long
    Dim now_teikan As Long, pop_obj As Object, pop_re As Long
    
    ThisWorkbook.Worksheets(w2pdata_sheet).Unprotect
    
    'パターン分けの色を定義
    patern_1_color = RGB(color1_R, color1_G, color1_B) '紫色
    patern_2_color = RGB(color2_R, color2_G, color2_B) '緑色
    patern_3_color = RGB(color3_R, color3_G, color3_B) '赤色
    patern_4_color = RGB(color4_R, color4_G, color4_B) '黄色
    
    Application.ScreenUpdating = False
    
    '「定款コード」シートの内容取得
    t_code_data.s_row = 2
    t_code_data.code_clm = 1
    t_code_data.out_title_row = 1
    t_code_data.flg = False
    With ThisWorkbook.Worksheets(teikan_code_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        If end_row = 1 Then
        Else
            end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
            t_code_data.list = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
            t_code_data.flg = True
        End If
    End With
        
    '--- W2P用 列マッピング設定（143列） ---
    with csv_data
        .title_row = 1
        .store_clm = 1
        .order_nom_clm = 2
        .order_status_clm = 19
        .item_code_clm = 21
        .item_name_clm = 22
        .item_count_clm = 23
        .haisou_order_name_clm = 8
        .haisou_name_clm = 12
        .haisou_post_clm = 13
        .haisou_address1_clm = 14
        .haisou_address2_clm = 15
        .haisou_address3_clm = 16
        .haisou_tel_clm = 17
        .haisousaki_tantou_clm = 18
    end with

    ' W2Pシート列マッピング（対応する方）
    with w2p_data
        .title_row = 1
        .store_clm = 1
        .order_nom_clm = 2
        .order_date_clm = 4
        .haisou_order_name_clm = 8
        .order_status_clm = 19
        .item_code_clm = 20
        .item_name_clm = 21
        .item_count_clm = 22
        .haisou_name_clm = 12
        .haisou_post_clm = 13
        .haisou_address1_clm = 14
        .haisou_address2_clm = 15
        .haisou_address3_clm = 16
        .haisou_tel_clm = 17
        .haisousaki_tantou_clm = 18
        .haisou_name_item_clm = 25
        .haisou_tantou_item_clm = 31
        .nouki_clm = 32
        .sagyou_shiji_clm = 33
        .syukko_yotei_clm = 34
        .end_data_clm = 34
    end with
        
    '「配送先住所」シート情報取得
    with haisousaki_data
        .patern_key_row = 1
        .title_row = 2
        .nomber_clm = 1
    End With
    '配送先住所データ取得               
    With ThisWorkbook.Worksheets(haisousaki_address_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
        haisousaki_data.address_data = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
    End With
    
    'ファイル名設定シート情報取得
    with set_file_name_data
        .start_row = 3
        .start_clm = 1
        .file_clm = 1
        .file_name_clm = 2
        .order_list_row = 1
        .shindou_list_row = 2
        .kyoten_list_row = 3
        .maru_list_row = 4
        .teikan_list_row = 7
        .end_row = 9
        .end_clm = 2
    End With

    'ファイル名リスト取得
    With ThisWorkbook.Worksheets(set_file_name_sheet)
        set_file_name_data.file_name_list = .Range(.Cells(set_file_name_data.start_row, set_file_name_data.start_clm), .Cells(set_file_name_data.end_row, set_file_name_data.end_clm)).Value
    End With
        
    '「【ヘッダー】出荷データ」シート情報取得
    With ThisWorkbook.Worksheets(hed_syukka_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
        hed_syukka = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
    End With
    
    '「【ヘッダー】入庫取込データ」シート情報取得
    With ThisWorkbook.Worksheets(hed_nyuuko_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
        hed_nyuuko = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
    End With
        
    ReDim w2p_data.w2p_list(1 To UBound(csv_data.csv_list, 1), 1 To w2p_data.end_data_clm)
        
        'CSVファイルの中身を「W2Pデータ貼り付け」シートの該当する列に貼り付け
        With ThisWorkbook.Worksheets(w2pdata_sheet)
            'タイトル保持
            For now_clm = 1 To w2p_data.end_data_clm
                w2p_data.w2p_list(1, now_clm) = .Cells(1, now_clm)
            Next
            
            'シートの初期化（2行目以降のデータ領域のみクリア）
            '今回書き込む行数（配列基準）
            Dim writeLastRow As Long
            writeLastRow = UBound(csv_data.csv_list, 1)
            'エクセルシートにある行数
            Dim excelLastRow As Long
            excelLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            Dim lastR As Long
            lastR = IIf(writeLastRow > excelLastRow, writeLastRow, excelLastRow)
            If lastR >= 2 Then
                .Range(.Cells(2, 1), .Cells(lastR, 34)).Clear  '値＋書式を消す
            End If

            '列の表示形式を設定
            .Columns(3).NumberFormatLocal = "0"  '明細番号（数値）
            .Columns(csv_data.haisou_tel_clm).NumberFormatLocal = "@"  '配送先電話番号（文字列）
            .Columns(20).NumberFormatLocal = "@"  '商品コード（文字列）
            
            '------------------------------------------------------------
            ' CSVファイルの中身を「W2Pデータ貼り付け」シートの該当する列に貼り付け（配列に詰める）
            '------------------------------------------------------------
            For now_row = 2 To UBound(csv_data.csv_list, 1)
            
                '--- ストア～注文状態まで貼り付け（CSV列1～19 → シート列1～19）---
                For now_clm = csv_data.store_clm To csv_data.order_status_clm
                    w2p_data.w2p_list(now_row, now_clm) = csv_data.csv_list(now_row, now_clm)
                Next
            
                '--- 商品情報の貼り付け ---
                w2p_data.w2p_list(now_row, 20) = csv_data.csv_list(now_row, 21) 'CSV列21:商品コード → シート列20
                w2p_data.w2p_list(now_row, 21) = csv_data.csv_list(now_row, 22) 'CSV列22:商品名 → シート列21
                w2p_data.w2p_list(now_row, 22) = csv_data.csv_list(now_row, 23) 'CSV列23:注文数量 → シート列22
                w2p_data.w2p_list(now_row, 23) = csv_data.csv_list(now_row, 24) 'CSV列24:単価 → シート列23
                w2p_data.w2p_list(now_row, 24) = csv_data.csv_list(now_row, 25) 'CSV列25:小計 → シート列24
            
                '--- アイテム別配送先情報の貼り付け（CSV列29～35 → シート列25～31）---
                '配送先名(ｱｲﾃﾑ別）から配送先担当者名(ｱｲﾃﾑ別）まで
                w2p_data.w2p_list(now_row, 25) = csv_data.csv_list(now_row, 29) 'CSV列29:配送先名(ｱｲﾃﾑ別） → シート列25
                w2p_data.w2p_list(now_row, 26) = csv_data.csv_list(now_row, 30) 'CSV列30:配送先郵便番号(ｱｲﾃﾑ別） → シート列26
                w2p_data.w2p_list(now_row, 27) = csv_data.csv_list(now_row, 31) 'CSV列31:配送先住所1(ｱｲﾃﾑ別） → シート列27
                w2p_data.w2p_list(now_row, 28) = csv_data.csv_list(now_row, 32) 'CSV列32:配送先住所2(ｱｲﾃﾑ別） → シート列28
                w2p_data.w2p_list(now_row, 29) = csv_data.csv_list(now_row, 33) 'CSV列33:配送先住所3(ｱｲﾃﾑ別） → シート列29
                w2p_data.w2p_list(now_row, 30) = csv_data.csv_list(now_row, 34) 'CSV列34:配送先電話番号(ｱｲﾃﾑ別） → シート列30
                w2p_data.w2p_list(now_row, 31) = csv_data.csv_list(now_row, 35) 'CSV列35:配送先担当者名(ｱｲﾃﾑ別） → シート列31
            
            Next now_row
            
            '--- 2行目以降のデータをシートに書き戻す ---
            rMax = UBound(w2p_data.w2p_list, 1)
            cMax = UBound(w2p_data.w2p_list, 2)
            
            If rMax >= 2 Then
                '配列の2行目以降を抽出して書き込み
                ReDim body(1 To rMax - 1, 1 To cMax)
                
                For r = 2 To rMax
                    For c = 1 To cMax
                        body(r - 1, c) = w2p_data.w2p_list(r, c)
                    Next c
                Next r
                
                '2行目から(rMax-1+1)行目まで書き込み = 2行目から rMax行目まで
                '※bodyは(1 To rMax-1)なので、書き込み範囲も rMax-1行分にする
                .Range(.Cells(2, 1), .Cells(rMax, cMax)).Value = body
            End If

            '列番号の対応付けの配列定義
            clm_link_data.title_row = 1
            clm_link_data.order_detail_row = 2
            clm_link_data.haisou_address_row = 3
            clm_link_data.w2p_data_row = 4
            clm_link_data.nyuuko_data_row = 5
            clm_link_data.syukka_data_row = 6
            
            ReDim clm_link_data.clm_link_list(1 To 6, 1 To UBound(csv_data.csv_list, 2))
            For now_clm = 1 To UBound(clm_link_data.clm_link_list, 2)
                clm_link_data.clm_link_list(1, now_clm) = csv_data.csv_list(1, now_clm)
            Next
            For now_clm = 1 To UBound(clm_link_data.clm_link_list, 2)
                'orderDetailの列番号格納
                For now_order_clm = 1 To UBound(csv_data.csv_list, 2)
                    If csv_data.csv_list(csv_data.title_row, now_order_clm) = clm_link_data.clm_link_list(clm_link_data.title_row, now_clm) Then
                        clm_link_data.clm_link_list(clm_link_data.order_detail_row, now_clm) = now_order_clm
                        Exit For
                    End If
                Next
                '配送先住所の列番号格納
                For now_address_clm = 1 To UBound(haisousaki_data.address_data, 2)
                    If haisousaki_data.address_data(haisousaki_data.title_row, now_address_clm) = clm_link_data.clm_link_list(clm_link_data.title_row, now_clm) Then
                        clm_link_data.clm_link_list(clm_link_data.haisou_address_row, now_clm) = now_address_clm
                        Exit For
                    End If
                Next
                'w2pシートの列番号格納
                For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                    If w2p_data.w2p_list(w2p_data.title_row, now_w2p_clm) = clm_link_data.clm_link_list(clm_link_data.title_row, now_clm) Then
                        clm_link_data.clm_link_list(clm_link_data.w2p_data_row, now_clm) = now_w2p_clm
                        Exit For
                    End If
                Next
            Next
            
            match_flg = False
            match_title_row = 1
            match_address_row = 2
            match_w2p_row = 3
            circle_count = 0
            
            '住所シートに○がついている列番号取得
            For now_key_clm = 1 To UBound(haisousaki_data.address_data, 2)
                If haisousaki_data.address_data(haisousaki_data.patern_key_row, now_key_clm) = "○" Then
                    circle_count = circle_count + 1
                    ReDim Preserve match_key(1 To 3, 1 To circle_count)
                    match_key(match_title_row, circle_count) = haisousaki_data.address_data(haisousaki_data.title_row, now_key_clm)
                    match_key(match_address_row, circle_count) = now_key_clm
                    For now_link_clm = 1 To UBound(clm_link_data.clm_link_list, 2)
                        If clm_link_data.clm_link_list(clm_link_data.haisou_address_row, now_link_clm) = now_key_clm Then
                            w2p_clm = clm_link_data.clm_link_list(clm_link_data.w2p_data_row, now_link_clm)
                        End If
                    Next
                    match_key(match_w2p_row, circle_count) = w2p_clm
                End If
            Next
            
            'パターン別色分け処理と納期、出庫予定日入力
            For now_row = 2 To UBound(w2p_data.w2p_list, 1)
                teikan_flg = False
                'データが入っている行に対して色分け処理を行う
                If w2p_data.w2p_list(now_row, w2p_data.item_code_clm) <> "" Then
                    send_time = get_send_time(w2p_data.w2p_list(now_row, w2p_data.haisou_address1_clm))
                    order_date = Date
                    match_flg = False
                                        
                    'まずは新藤様商品について、ユーザー指定の色分け条件に合致するかチェック
                    Set ws = ThisWorkbook.Worksheets(set_syouhin_code_sheet)
                    cs_data.clm_min_code = 1
                    cs_data.clm_max_code = 2
                    cs_data.clm_color = 3
                    cs_data.clm_to_send_days = 4
                    cs_data.start_row = 3
                    cs_data.end_row = GetLastRow(ws, cs_data.clm_min_code)
                    '[特定新藤様商品コード設定シート]の情報取得 & チェック
                    With ws
                        cs_data.list = .Range(.Cells(1, 1), .Cells(cs_data.end_row, cs_data.clm_to_send_days)).Value
                    End With
                    For now_cs_data_row = cs_data.start_row To UBound(cs_data.list)
                        If (cs_data.list(now_cs_data_row, cs_data.clm_min_code) <> "" And _
                            IsNumeric(cs_data.list(now_cs_data_row, cs_data.clm_min_code)) = True) And _
                           (cs_data.list(now_cs_data_row, cs_data.clm_max_code) <> "" And _
                            IsNumeric(cs_data.list(now_cs_data_row, cs_data.clm_max_code)) = True) And _
                           (cs_data.list(now_cs_data_row, cs_data.clm_to_send_days) <> "" And _
                            IsNumeric(cs_data.list(now_cs_data_row, cs_data.clm_to_send_days)) = True) Then
                            'Nothing
                        Else
                            GoTo ERR_END
                        End If
                    Next
                    '現在商品が特定新藤様商品コードかチェック
                    user_set_flg = False
                    For now_cs_data_row = cs_data.start_row To UBound(cs_data.list)
                        '商品コード種類が新藤様商品か(数字のみコードか)チェック
                        If IsNumeric(w2p_data.w2p_list(now_row, w2p_data.item_code_clm)) = True Then
                            now_syouhin_code = w2p_data.w2p_list(now_row, w2p_data.item_code_clm)
                            cs_data.now_min_code = cs_data.list(now_cs_data_row, cs_data.clm_min_code)
                            cs_data.now_max_code = cs_data.list(now_cs_data_row, cs_data.clm_max_code)
                            If now_syouhin_code >= cs_data.now_min_code And now_syouhin_code <= cs_data.now_max_code Then
                                '指定の商品コード範囲内だった場合
                                .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = ws.Cells(now_cs_data_row, cs_data.clm_color).Interior.Color
                                to_send_days = cs_data.list(now_cs_data_row, cs_data.clm_to_send_days)
                                retern_day = createDay(order_date, send_time, key_numeric, send_sindoh, to_send_days)
                                w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                                w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                                user_set_flg = True
                                Exit For
                            End If
                        End If
                    Next
    
                    'ユーザー指定の色分け条件に合致しなかった場合
                    If user_set_flg = False Then
                        '商品コードが「定款コード」に一致する場合
                        If t_code_data.flg = True Then
                            For now_teikan = t_code_data.s_row To UBound(t_code_data.list)
                                If w2p_data.w2p_list(now_row, w2p_data.item_code_clm) = CStr(t_code_data.list(now_teikan, t_code_data.code_clm)) Then
                                    teikan_flg = True
                                    Exit For
                                End If
                            Next
                        End If
                        If teikan_flg = True Then
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_4_color
                            '納期、出荷予定日の入力はしない
                        
                        '商品コードが数字のみの場合
                        ElseIf IsNumeric(w2p_data.w2p_list(now_row, w2p_data.item_code_clm)) = True Then
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_1_color
                            retern_day = createDay(order_date, send_time, key_numeric, send_sindoh)
                            w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                            w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                            
                        '商品コードがAから始まっていた場合
                        ElseIf Left(w2p_data.w2p_list(now_row, w2p_data.item_code_clm), 1) = "A" Then
                            For now_haisou_row = 3 To UBound(haisousaki_data.address_data, 1)
                                match_flg = True
                                For match_key_clm = 1 To UBound(match_key, 2)
                                    If w2p_data.w2p_list(now_row, match_key(match_w2p_row, match_key_clm)) <> haisousaki_data.address_data(now_haisou_row, match_key(match_address_row, match_key_clm)) Then
                                        match_flg = False
                                        Exit For
                                    End If
                                Next
                                If match_flg = True Then
                                    If haisousaki_data.address_data(now_haisou_row, haisousaki_data.nomber_clm) <> kyoten_nom Then
                                        Exit For
                                    Else
                                        match_flg = False
                                    End If
                                End If
                            Next
                            If match_flg = False Then
                                .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_2_color
                            Else
                                .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_3_color
                                retern_day = createDay(order_date, send_time, key_A, send_sompo)
                                w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                                w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                            End If
                        
                        '商品コードがBから始まっていた場合
                        ElseIf Left(w2p_data.w2p_list(now_row, w2p_data.item_code_clm), 1) = "B" Then
                            For now_haisou_row = 3 To UBound(haisousaki_data.address_data, 1)
                                match_flg = True
                                For match_key_clm = 1 To UBound(match_key, 2)
                                    If w2p_data.w2p_list(now_row, match_key(match_w2p_row, match_key_clm)) <> haisousaki_data.address_data(now_haisou_row, match_key(match_address_row, match_key_clm)) Then
                                        match_flg = False
                                        Exit For
                                    End If
                                Next
                                If match_flg = True Then
                                    If haisousaki_data.address_data(now_haisou_row, haisousaki_data.nomber_clm) = honsya_nom Then
                                        Exit For
                                    Else
                                        match_flg = False
                                    End If
                                End If
                            Next
                            If match_flg = False Then
                                .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_2_color
                            Else
                                .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_3_color
                                retern_day = createDay(order_date, send_time, key_B, send_sompo)
                                w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                                w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                            End If
        
                        '商品コードがCから始まっていた場合
                        ElseIf Left(w2p_data.w2p_list(now_row, w2p_data.item_code_clm), 1) = "C" Then
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_3_color
                            retern_day = createDay(order_date, send_time, key_C, send_sompo)
                            w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                            w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                        End If
                    End If
                End If
            Next
            
        '--- 色分け処理後、納期・作業指示書・出庫予定日を3列まとめて一括書き戻し（2行目から） ---
        Dim lastRow As Long, rCount As Long
        Dim outArr() As Variant
        Dim i As Long, srcRow As Long
        lastRow = UBound(w2p_data.w2p_list, 1)
        If lastRow < 2 Then Exit Sub
        rCount = lastRow - 1              '2行目～最終行の行数
        ReDim outArr(1 To rCount, 1 To 3) '（納期, 作業指示, 出庫予定）
        For srcRow = 2 To lastRow
            i = srcRow - 1
            outArr(i, 1) = w2p_data.w2p_list(srcRow, w2p_data.nouki_clm)
            outArr(i, 2) = w2p_data.w2p_list(srcRow, w2p_data.sagyou_shiji_clm)
            outArr(i, 3) = w2p_data.w2p_list(srcRow, w2p_data.syukko_yotei_clm)
        Next srcRow
        .Range(.Cells(2, w2p_data.nouki_clm), .Cells(lastRow, w2p_data.syukko_yotei_clm)).Value2 = outArr
   
                  
        End With
        
        ThisWorkbook.Worksheets(w2pdata_sheet).Activate
        
        Application.ScreenUpdating = True
        
        Set pop_obj = CreateObject("WScript.Shell")
        pop_re = pop_obj.Popup("色分けされたパターンは以下の通りです。" & vbCrLf & vbCrLf & _
        "新藤Cに手配依頼するデータ：紫色" & vbCrLf & _
        "マルテックスで商品をピックし、RLCが配送するデータ：緑色" & vbCrLf & _
        "マルテックスが配送まで手配するデータ：赤色" & vbCrLf & _
        "定款：黄色" & vbCrLf & _
        "特定新藤様商品コードのデータ：ユーザー指定の色" & vbCrLf & vbCrLf & _
        "確認後、「作業指示書作成」ボタンを押下してください。", 0, "確認", vbOKOnly)
        
        '「w2pデータ貼り付け」シート保護
        ThisWorkbook.Worksheets(w2pdata_sheet).Protect AllowFiltering:=True
        
        Exit Sub
    
ERR_END:
    MsgBox ("[特定新藤様商品コード設定シート]に不正があります。" & vbCrLf & _
            "以下のような不正がないか、確認してください。" & vbCrLf & _
            " ・未記入列が存在する" & vbCrLf & _
            " ・新藤様商品以外のコード(アルファベット有コード)が記入されている" & vbCrLf & _
            " ・営業日数に数値以外が含まれる")

End Sub


'==========================================================
' Spinno処理プロシージャ（20列CSV用）
' 機能: Spinno形式のCSVデータ（20列）を読み込み、
'       W2Pシート形式に変換・色分け・納期計算
'==========================================================
Sub process_spinno(csv_data As csv_data)
    
    '変数宣言
    Dim patern_1_color As Long, patern_2_color As Long
    Dim patern_3_color As Long, patern_4_color As Long
    Dim w2p_data As w2p_data
    Dim t_code_data As teikan_code_data
    Dim haisousaki_data As haisousaki_data
    Dim set_file_name_data As set_file_name_data
    Dim hed_syukka() As Variant, hed_nyuuko() As Variant
    Dim clm_link_data As clm_link_data
    Dim end_row As Long, end_clm As Long
    Dim now_row As Long, now_clm As Long
    Dim lastUsedRow As Long
    Dim rMax As Long, cMax As Long
    Dim body() As Variant
    Dim r As Long, c As Long
    Dim now_order_clm As Long, now_address_clm As Long, now_w2p_clm As Long
    Dim order_internal_id_clm As Long, customer_company_clm As Long
    Dim group_name_clm As Long, user_number_clm As Long
    Dim match_flg As Boolean
    Dim match_key() As Variant
    Dim match_title_row As Long, match_address_row As Long, match_w2p_row As Long
    Dim circle_count As Long, now_key_clm As Long, now_link_clm As Long, w2p_clm As Long
    Dim RET_DAY As RET_DAY, send_time As Long, order_date As Date
    Dim to_send_days As Long, retern_day As RET_DAY, teikan_flg As Boolean
    Dim cs_data As color_send_days_data, ws As Worksheet
    Dim user_set_flg As Boolean, now_syouhin_code As Long, key As String
    Dim now_cs_data_row As Long, now_haisou_row As Long, match_key_clm As Long
    Dim now_teikan As Long, pop_obj As Object, pop_re As Long
    
    ThisWorkbook.Worksheets(w2pdata_sheet).Unprotect
    
    'パターン分けの色を定義
    patern_1_color = RGB(color1_R, color1_G, color1_B) '紫色
    patern_2_color = RGB(color2_R, color2_G, color2_B) '緑色
    patern_3_color = RGB(color3_R, color3_G, color3_B) '赤色
    patern_4_color = RGB(color4_R, color4_G, color4_B) '黄色
    
    Application.ScreenUpdating = False
    
    '「定款コード」シートの内容取得
    t_code_data.s_row = 2
    t_code_data.code_clm = 1
    t_code_data.out_title_row = 1
    t_code_data.flg = False
    With ThisWorkbook.Worksheets(teikan_code_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        If end_row = 1 Then
        Else
            end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
            t_code_data.list = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
            t_code_data.flg = True
        End If
    End With
        
    '--- Spinno用 列マッピング設定（20列） ---
    csv_data.title_row = 1
    csv_data.store_clm = 1 'A 注文者タイプ
    csv_data.order_nom_clm = 2 'B 注文番号
    csv_data.order_date_clm = 4 'D 発注依頼日
    csv_data.haisou_order_name_clm = 8 'H 注文者氏名
    csv_data.haisou_name_clm = 11  'K 配送先会社名
    csv_data.haisou_post_clm = 12  'L 配送先郵便番号
    csv_data.haisou_address1_clm = 13 'M 配送先都道府県
    csv_data.haisou_address2_clm = 14 'N 配送先住所1
    csv_data.haisou_address3_clm = 15 'O 配送先住所2
    csv_data.haisou_tel_clm = 16 'P 配送先電話番号
    csv_data.order_status_clm = 17 'Q ステータス
    csv_data.item_code_clm = 18 'R アイテムコード
    csv_data.item_name_clm = 19 'S アイテム名
    csv_data.item_count_clm = 20 'T 明細別数量
    
    'Spinno固有の列（既に先頭で宣言済み）
    order_internal_id_clm = 3  'C 注文内部管理番号
    customer_company_clm = 5  'E 注文者会社名
    group_name_clm = 6  'F グループ名
    user_number_clm = 7  'G ユーザー番号
    csv_data.item_code_clm = 18 'R アイテムコード
    csv_data.item_name_clm = 19 'S アイテム名
    csv_data.item_count_clm = 20 'T 明細別数量

    ' W2Pシート列マッピング
    w2p_data.title_row = 1
    w2p_data.store_clm = 1 'A ストア
    w2p_data.order_nom_clm = 2 'C 発注番号
    w2p_data.order_date_clm = 4 'E 注文日
    w2p_data.haisou_order_name_clm = 8 'I 発注者
    w2p_data.haisou_name_clm = 12 'L 配送先名
    w2p_data.haisou_post_clm = 13 'M 配送先郵便番号
    w2p_data.haisou_address1_clm = 14 'N 配送先住所1
    w2p_data.haisou_address2_clm = 15 'O 配送先住所2
    w2p_data.haisou_address3_clm = 16 'P 配送先住所3
    w2p_data.haisou_tel_clm = 17 'Q 配送先電話番号
    w2p_data.order_status_clm = 19 'R 注文状態
    w2p_data.item_code_clm = 20 'S 商品コード
    w2p_data.item_name_clm = 21 'T 商品名
    w2p_data.item_count_clm = 22 'U 注文数量
    w2p_data.nouki_clm = 32  'AF 納期
    w2p_data.sagyou_shiji_clm = 33 'AG 作業指示書
    w2p_data.syukko_yotei_clm = 34  'AH 出庫予定日
    w2p_data.end_data_clm = 34 '最終列
        
    '「配送先住所」シート情報取得
    haisousaki_data.patern_key_row = 1
    haisousaki_data.title_row = 2
    haisousaki_data.nomber_clm = 1
                            
    With ThisWorkbook.Worksheets(haisousaki_address_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
        haisousaki_data.address_data = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
    End With
    
    'ファイル名設定シート情報取得
    set_file_name_data.start_row = 3
    set_file_name_data.start_clm = 1
    set_file_name_data.file_clm = 1
    set_file_name_data.file_name_clm = 2
    set_file_name_data.order_list_row = 1
    set_file_name_data.shindou_list_row = 2
    set_file_name_data.kyoten_list_row = 3
    set_file_name_data.maru_list_row = 4
    set_file_name_data.teikan_list_row = 7
    set_file_name_data.end_row = 9
    set_file_name_data.end_clm = 2
    With ThisWorkbook.Worksheets(set_file_name_sheet)
        set_file_name_data.file_name_list = .Range(.Cells(set_file_name_data.start_row, set_file_name_data.start_clm), .Cells(set_file_name_data.end_row, set_file_name_data.end_clm)).Value
    End With
    
    '「【ヘッダー】出荷データ」シート情報取得
    With ThisWorkbook.Worksheets(hed_syukka_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
        hed_syukka = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
    End With
    
    '「【ヘッダー】入庫取込データ」シート情報取得
    With ThisWorkbook.Worksheets(hed_nyuuko_sheet)
        end_row = .Cells(.Rows.count, 1).End(xlUp).row
        end_clm = .Cells(1, .Columns.count).End(xlToLeft).Column
        hed_nyuuko = .Range(.Cells(1, 1), .Cells(end_row, end_clm)).Value
    End With
    
    ReDim w2p_data.w2p_list(1 To UBound(csv_data.csv_list, 1), 1 To w2p_data.end_data_clm)
    
    'CSVファイルの中身を「W2Pデータ貼り付け」シートの該当する列に貼り付け
    With ThisWorkbook.Worksheets(w2pdata_sheet)
        'タイトル保持
        For now_clm = 1 To w2p_data.end_data_clm
            w2p_data.w2p_list(1, now_clm) = .Cells(1, now_clm)
        Next
        
        'シートの初期化（2行目以降のデータ領域のみクリア）
        '今回書き込む行数（配列基準）
        Dim writeLastRow As Long
        writeLastRow = UBound(csv_data.csv_list, 1)
        'エクセルシートにある行数
        Dim excelLastRow As Long
        excelLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Dim lastR As Long
        lastR = IIf(writeLastRow > excelLastRow, writeLastRow, excelLastRow)
        If lastR >= 2 Then
            .Range(.Cells(2, 1), .Cells(lastR, 34)).Clear  '値＋書式を消す
        End If

        '列の表示形式を設定
        .Columns(3).NumberFormatLocal = "0"  '明細番号（数値）
        .Columns(csv_data.haisou_tel_clm).NumberFormatLocal = "@"  '配送先電話番号（文字列）
        
        '------------------------------------------------------------
        ' Spinno CSVファイルの中身を「W2Pデータ貼り付け」シートの該当する列に貼り付け（配列に詰める）
        '------------------------------------------------------------
        For now_row = 2 To UBound(csv_data.csv_list, 1)
        
            '--- ストア（CSV列1） → W2Pシート列1（「SOMPOケア　」を付加） ---
            If csv_data.store_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.store_clm) = "SOMPOケア　" & csv_data.csv_list(now_row, csv_data.store_clm)
            End If
            
            '--- 注文番号（CSV列2） → W2Pシート列2 ---
            If csv_data.order_nom_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.order_nom_clm) = csv_data.csv_list(now_row, csv_data.order_nom_clm)
            End If
            
            '--- 注文内部管理番号（CSV列3） → W2Pシート列3（明細番号） ---
            If order_internal_id_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, 3) = csv_data.csv_list(now_row, order_internal_id_clm)
            End If
            
            '--- 発注依頼日（CSV列4） → W2Pシート列4（注文日） ---
            If csv_data.order_date_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.order_date_clm) = csv_data.csv_list(now_row, csv_data.order_date_clm)
            End If
            
            '--- 注文者会社名（CSV列5） → W2Pシート列5（顧客） ---
            If customer_company_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, 5) = csv_data.csv_list(now_row, customer_company_clm)
            End If
            
            '--- グループ名（CSV列6） → W2Pシート列6（グループ） ---
            If group_name_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, 6) = csv_data.csv_list(now_row, group_name_clm)
            End If
            
            '--- ユーザー番号（CSV列7） → W2Pシート列7（発注者ID） ---
            If user_number_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, 7) = csv_data.csv_list(now_row, user_number_clm)
            End If
            
            '--- 注文者氏名（CSV列8） → W2Pシート列8（発注者） ---
            If csv_data.haisou_order_name_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.haisou_order_name_clm) = csv_data.csv_list(now_row, csv_data.haisou_order_name_clm)
            End If
            
            '--- CSV列9（メール/ログインID） → W2Pシート列9（発注者ログインID） ---
            If 9 <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, 9) = csv_data.csv_list(now_row, 9)
            End If

            '--- CSV列10（発注者コード） → W2Pシート列10（発注者コード） ---
            If 10 <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, 10) = csv_data.csv_list(now_row, 10)
            End If
            
            '--- 配送先会社名（CSV列11） → W2Pシート列12（配送先名） ---
            If csv_data.haisou_name_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.haisou_name_clm) = csv_data.csv_list(now_row, csv_data.haisou_name_clm)
            End If
            
            '--- 配送先郵便番号（CSV列12） → W2Pシート列13 ---
            If csv_data.haisou_post_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.haisou_post_clm) = csv_data.csv_list(now_row, csv_data.haisou_post_clm)
            End If
            
            '--- 配送先都道府県（CSV列13） → W2Pシート列14（配送先住所1） ---
            If csv_data.haisou_address1_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.haisou_address1_clm) = csv_data.csv_list(now_row, csv_data.haisou_address1_clm)
            End If
            
            '--- 配送先住所1（CSV列14） → W2Pシート列15（配送先住所2） ---
            If csv_data.haisou_address2_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.haisou_address2_clm) = csv_data.csv_list(now_row, csv_data.haisou_address2_clm)
            End If
            
            '--- 配送先住所2（CSV列15） → W2Pシート列16（配送先住所3） ---
            If csv_data.haisou_address3_clm <= UBound(csv_data.csv_list, 2) Then
                Dim tmpAddr3 As String
                tmpAddr3 = CStr(csv_data.csv_list(now_row, csv_data.haisou_address3_clm))
                ' 文字化けしている全角疑問符や半角?をハイフンに変換
                tmpAddr3 = Replace(tmpAddr3, "？", "-")
                tmpAddr3 = Replace(tmpAddr3, "?", "-")
                w2p_data.w2p_list(now_row, w2p_data.haisou_address3_clm) = tmpAddr3
            End If
            ' Spinno同様、住所2と住所3を結合して住所2に格納、住所3は空にする
            On Error Resume Next
            If w2p_data.haisou_address2_clm >= 1 And w2p_data.haisou_address3_clm >= 1 Then
                Dim a2 As String, a3 As String
                a2 = CStr(w2p_data.w2p_list(now_row, w2p_data.haisou_address2_clm))
                a3 = CStr(w2p_data.w2p_list(now_row, w2p_data.haisou_address3_clm))
                If Trim(a3) <> "" Then
                    w2p_data.w2p_list(now_row, w2p_data.haisou_address2_clm) = Trim(a2 & " " & a3)
                    w2p_data.w2p_list(now_row, w2p_data.haisou_address3_clm) = ""
                End If
            End If
            On Error GoTo 0
            
            '--- 配送先電話番号（CSV列16） → W2Pシート列17 ---
            If csv_data.haisou_tel_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.haisou_tel_clm) = csv_data.csv_list(now_row, csv_data.haisou_tel_clm)
            End If
            
            '--- ステータス（CSV列17） → W2Pシート列19（注文状態） ---
            If csv_data.order_status_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.order_status_clm) = csv_data.csv_list(now_row, csv_data.order_status_clm)
            End If
            
            '--- アイテムコード（CSV列18） → W2Pシート列20（商品コード） ---
            If csv_data.item_code_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.item_code_clm) = csv_data.csv_list(now_row, csv_data.item_code_clm)
            End If
            
            '--- アイテム名（CSV列19） → W2Pシート列21（商品名） ---
            If csv_data.item_name_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.item_name_clm) = csv_data.csv_list(now_row, csv_data.item_name_clm)
            End If
            
            '--- 明細別数量（CSV列20） → W2Pシート列22（注文数量） ---
            If csv_data.item_count_clm <= UBound(csv_data.csv_list, 2) Then
                w2p_data.w2p_list(now_row, w2p_data.item_count_clm) = csv_data.csv_list(now_row, csv_data.item_count_clm)
            End If
        
        Next now_row
'---ここまでは重くない------------------------------------------------------------------------
        
        '--- 1行目（タイトル）を触らず、2行目以降だけ書き戻す ---
        rMax = UBound(w2p_data.w2p_list, 1)
        cMax = UBound(w2p_data.w2p_list, 2)
        
        If rMax >= 2 Then
            ReDim body(1 To rMax - 1, 1 To cMax)
        
            For r = 2 To rMax
                For c = 1 To cMax
                    body(r - 1, c) = w2p_data.w2p_list(r, c)
                Next c
            Next r
        
            .Range(.Cells(2, 1), .Cells(rMax, cMax)).Value = body
        End If

        
        '列番号の対応付けの配列定義
        clm_link_data.title_row = 1
        clm_link_data.order_detail_row = 2
        clm_link_data.haisou_address_row = 3
        clm_link_data.w2p_data_row = 4
        clm_link_data.nyuuko_data_row = 5
        clm_link_data.syukka_data_row = 6
        
        ReDim clm_link_data.clm_link_list(1 To 6, 1 To UBound(csv_data.csv_list, 2))
        For now_clm = 1 To UBound(clm_link_data.clm_link_list, 2)
            clm_link_data.clm_link_list(1, now_clm) = csv_data.csv_list(1, now_clm)
        Next
        For now_clm = 1 To UBound(clm_link_data.clm_link_list, 2)
            'orderDetailの列番号格納
            For now_order_clm = 1 To UBound(csv_data.csv_list, 2)
                If csv_data.csv_list(csv_data.title_row, now_order_clm) = clm_link_data.clm_link_list(clm_link_data.title_row, now_clm) Then
                    clm_link_data.clm_link_list(clm_link_data.order_detail_row, now_clm) = now_order_clm
                    Exit For
                End If
            Next
            '配送先住所の列番号格納
            For now_address_clm = 1 To UBound(haisousaki_data.address_data, 2)
                If haisousaki_data.address_data(haisousaki_data.title_row, now_address_clm) = clm_link_data.clm_link_list(clm_link_data.title_row, now_clm) Then
                    clm_link_data.clm_link_list(clm_link_data.haisou_address_row, now_clm) = now_address_clm
                    Exit For
                End If
            Next
            'w2pシートの列番号格納
            For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                If w2p_data.w2p_list(w2p_data.title_row, now_w2p_clm) = clm_link_data.clm_link_list(clm_link_data.title_row, now_clm) Then
                    clm_link_data.clm_link_list(clm_link_data.w2p_data_row, now_clm) = now_w2p_clm
                    Exit For
                End If
            Next
        Next
        
        match_flg = False
        match_title_row = 1
        match_address_row = 2
        match_w2p_row = 3
        circle_count = 0
        
        '住所シートに○がついている列番号取得
        For now_key_clm = 1 To UBound(haisousaki_data.address_data, 2)
            If haisousaki_data.address_data(haisousaki_data.patern_key_row, now_key_clm) = "○" Then
                circle_count = circle_count + 1
                ReDim Preserve match_key(1 To 3, 1 To circle_count)
                match_key(match_title_row, circle_count) = haisousaki_data.address_data(haisousaki_data.title_row, now_key_clm)
                match_key(match_address_row, circle_count) = now_key_clm
                For now_link_clm = 1 To UBound(clm_link_data.clm_link_list, 2)
                    If clm_link_data.clm_link_list(clm_link_data.haisou_address_row, now_link_clm) = now_key_clm Then
                        w2p_clm = clm_link_data.clm_link_list(clm_link_data.w2p_data_row, now_link_clm)
                    End If
                Next
                match_key(match_w2p_row, circle_count) = w2p_clm
            End If
        Next
'---ここは少し重い（でも要因違いそう）----------------------------------------------------------------
        'パターン別色分け処理と納期、出庫予定日入力
        For now_row = 2 To UBound(w2p_data.w2p_list, 1)
            teikan_flg = False
            'データが入っている行に対して色分け処理を行う
            If w2p_data.w2p_list(now_row, w2p_data.item_code_clm) <> "" Then
                send_time = get_send_time(w2p_data.w2p_list(now_row, w2p_data.haisou_address1_clm))
                order_date = Date
                match_flg = False
                                    
                'まずは新藤様商品について、ユーザー指定の色分け条件に合致するかチェック
                Set ws = ThisWorkbook.Worksheets(set_syouhin_code_sheet)
                cs_data.clm_min_code = 1
                cs_data.clm_max_code = 2
                cs_data.clm_color = 3
                cs_data.clm_to_send_days = 4
                cs_data.start_row = 3
                cs_data.end_row = GetLastRow(ws, cs_data.clm_min_code)
                '[特定新藤様商品コード設定シート]の情報取得 & チェック
                With ws
                    cs_data.list = .Range(.Cells(1, 1), .Cells(cs_data.end_row, cs_data.clm_to_send_days)).Value
                End With
                For now_cs_data_row = cs_data.start_row To UBound(cs_data.list)
                    If (cs_data.list(now_cs_data_row, cs_data.clm_min_code) <> "" And _
                        IsNumeric(cs_data.list(now_cs_data_row, cs_data.clm_min_code)) = True) And _
                       (cs_data.list(now_cs_data_row, cs_data.clm_max_code) <> "" And _
                        IsNumeric(cs_data.list(now_cs_data_row, cs_data.clm_max_code)) = True) And _
                       (cs_data.list(now_cs_data_row, cs_data.clm_to_send_days) <> "" And _
                        IsNumeric(cs_data.list(now_cs_data_row, cs_data.clm_to_send_days)) = True) Then
                        'Nothing
                    Else
                        GoTo ERR_END_SPINNO
                    End If
                Next
                '現在商品が特定新藤様商品コードかチェック
                user_set_flg = False
                For now_cs_data_row = cs_data.start_row To UBound(cs_data.list)
                    '商品コード種類が新藤様商品か(数字のみコードか)チェック
                    If IsNumeric(w2p_data.w2p_list(now_row, w2p_data.item_code_clm)) = True Then
                        now_syouhin_code = w2p_data.w2p_list(now_row, w2p_data.item_code_clm)
                        cs_data.now_min_code = cs_data.list(now_cs_data_row, cs_data.clm_min_code)
                        cs_data.now_max_code = cs_data.list(now_cs_data_row, cs_data.clm_max_code)
                        If now_syouhin_code >= cs_data.now_min_code And now_syouhin_code <= cs_data.now_max_code Then
                            '指定の商品コード範囲内だった場合
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = ws.Cells(now_cs_data_row, cs_data.clm_color).Interior.Color
                            to_send_days = cs_data.list(now_cs_data_row, cs_data.clm_to_send_days)
                            retern_day = createDay(order_date, send_time, key_numeric, send_sindoh, to_send_days)
                            w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                            w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                            user_set_flg = True
                            Exit For
                        End If
                    End If
                Next

                'ユーザー指定の色分け条件に合致しなかった場合
                If user_set_flg = False Then
                    '商品コードが「定款コード」に一致する場合
                    If t_code_data.flg = True Then
                        For now_teikan = t_code_data.s_row To UBound(t_code_data.list)
                            If w2p_data.w2p_list(now_row, w2p_data.item_code_clm) = CStr(t_code_data.list(now_teikan, t_code_data.code_clm)) Then
                                teikan_flg = True
                                Exit For
                            End If
                        Next
                    End If
                    If teikan_flg = True Then
                        .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_4_color
                        '納期、出荷予定日の入力はしない
                    
                    '商品コードが数字のみの場合
                    ElseIf IsNumeric(w2p_data.w2p_list(now_row, w2p_data.item_code_clm)) = True Then
                        .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_1_color
                        retern_day = createDay(order_date, send_time, key_numeric, send_sindoh)
                        w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                        w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                        
                    '商品コードがAから始まっていた場合
                    ElseIf Left(w2p_data.w2p_list(now_row, w2p_data.item_code_clm), 1) = "A" Then
                        For now_haisou_row = 3 To UBound(haisousaki_data.address_data, 1)
                            match_flg = True
                            For match_key_clm = 1 To UBound(match_key, 2)
                                If w2p_data.w2p_list(now_row, match_key(match_w2p_row, match_key_clm)) <> haisousaki_data.address_data(now_haisou_row, match_key(match_address_row, match_key_clm)) Then
                                    match_flg = False
                                    Exit For
                                End If
                            Next
                            If match_flg = True Then
                                If haisousaki_data.address_data(now_haisou_row, haisousaki_data.nomber_clm) <> kyoten_nom Then
                                    Exit For
                                Else
                                    match_flg = False
                                End If
                            End If
                        Next
                        If match_flg = False Then
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_2_color
                        Else
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_3_color
                            retern_day = createDay(order_date, send_time, key_A, send_sompo)
                            w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                            w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                        End If
                    
                    '商品コードがBから始まっていた場合
                    ElseIf Left(w2p_data.w2p_list(now_row, w2p_data.item_code_clm), 1) = "B" Then
                        For now_haisou_row = 3 To UBound(haisousaki_data.address_data, 1)
                            match_flg = True
                            For match_key_clm = 1 To UBound(match_key, 2)
                                If w2p_data.w2p_list(now_row, match_key(match_w2p_row, match_key_clm)) <> haisousaki_data.address_data(now_haisou_row, match_key(match_address_row, match_key_clm)) Then
                                    match_flg = False
                                    Exit For
                                End If
                            Next
                            If match_flg = True Then
                                If haisousaki_data.address_data(now_haisou_row, haisousaki_data.nomber_clm) = honsya_nom Then
                                    Exit For
                                Else
                                    match_flg = False
                                End If
                            End If
                        Next
                        If match_flg = False Then
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_2_color
                        Else
                            .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_3_color
                            retern_day = createDay(order_date, send_time, key_B, send_sompo)
                            w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                            w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                        End If
    
                    '商品コードがCから始まっていた場合
                    ElseIf Left(w2p_data.w2p_list(now_row, w2p_data.item_code_clm), 1) = "C" Then
                        .Range(.Cells(now_row, 1), .Cells(now_row, w2p_data.end_data_clm)).Interior.Color = patern_3_color
                        retern_day = createDay(order_date, send_time, key_C, send_sompo)
                        w2p_data.w2p_list(now_row, w2p_data.nouki_clm) = Format(retern_day.get_day, "YYYYMMDD")
                        w2p_data.w2p_list(now_row, w2p_data.syukko_yotei_clm) = Format(retern_day.send_day, "YYYYMMDD")
                    End If
                End If
            End If
        Next
        
        '--- 色分け処理後、納期・作業指示書・出庫予定日を3列まとめて一括書き戻し（2行目から） ---
        Dim lastRow As Long, rCount As Long
        Dim outArr() As Variant
        Dim i As Long, srcRow As Long
        lastRow = UBound(w2p_data.w2p_list, 1)
        If lastRow < 2 Then Exit Sub
        rCount = lastRow - 1              '2行目～最終行の行数
        ReDim outArr(1 To rCount, 1 To 3) '（納期, 作業指示, 出庫予定）
        For srcRow = 2 To lastRow
            i = srcRow - 1
            outArr(i, 1) = w2p_data.w2p_list(srcRow, w2p_data.nouki_clm)
            outArr(i, 2) = w2p_data.w2p_list(srcRow, w2p_data.sagyou_shiji_clm)
            outArr(i, 3) = w2p_data.w2p_list(srcRow, w2p_data.syukko_yotei_clm)
        Next srcRow
        .Range(.Cells(2, w2p_data.nouki_clm), .Cells(lastRow, w2p_data.syukko_yotei_clm)).Value2 = outArr
   
    End With
    
    ThisWorkbook.Worksheets(w2pdata_sheet).Activate
    
    Application.ScreenUpdating = True
    
    Set pop_obj = CreateObject("WScript.Shell")
    pop_re = pop_obj.Popup("色分けされたパターンは以下の通りです。" & vbCrLf & vbCrLf & _
    "新藤Cに手配依頼するデータ：紫色" & vbCrLf & _
    "マルテックスで商品をピックし、RLCが配送するデータ：緑色" & vbCrLf & _
    "マルテックスが配送まで手配するデータ：赤色" & vbCrLf & _
    "定款：黄色" & vbCrLf & _
    "特定新藤様商品コードのデータ：ユーザー指定の色" & vbCrLf & vbCrLf & _
    "確認後、「作業指示書作成」ボタンを押下してください。", 0, "確認", vbOKOnly)
    
    '「w2pデータ貼り付け」シート保護
    ThisWorkbook.Worksheets(w2pdata_sheet).Protect AllowFiltering:=True
    
    Exit Sub

ERR_END_SPINNO:
    MsgBox ("[特定新藤様商品コード設定シート]に不正があります。" & vbCrLf & _
            "以下のような不正がないか、確認してください。" & vbCrLf & _
            " ・未記入列が存在する" & vbCrLf & _
            " ・新藤様商品以外のコード(アルファベット有コード)が記入されている" & vbCrLf & _
            " ・営業日数に数値以外が含まれる")

End Sub
        

Function splitCsv(ByVal sp_str As String) As Variant
    'CSVの1行を指定したとき、その行の内容を判別して区切る
    
    '変数宣言
    Dim word() As Variant
    Dim rep_str As String
    Dim flg_str As Boolean
    Dim idx_chr As Long
    Dim pos_start As Long
    Dim str_chr As String
    Dim count_dq As Long
    Dim idx_wd As Long
    
    '前方から1文字ずつダブルクォーテーションを確認する
    rep_str = sp_str
    flg_str = False
    
    ReDim word(1 To 1)
    
    idx_chr = 1
    pos_start = 1
    Do While idx_chr <= Len(sp_str)
        str_chr = Mid(sp_str, idx_chr, 1)
        
        If str_chr = """" Then
            count_dq = doubleQuatCount(sp_str, idx_chr)
        
            If count_dq Mod 2 = 1 Then
                '奇数の場合は、文字列の開始または終了であるため、フラグを設定する
                If flg_str = True Then
                    flg_str = False
                Else
                    flg_str = True
                End If
                '端数(奇数分)のダブルクォーテーションを破棄
                sp_str = Left(sp_str, idx_chr - 1) & Right(sp_str, Len(sp_str) - idx_chr)
                count_dq = count_dq - 1
            End If
            'エスケープとしてダブルクォートの数を半分に減らし、その分だけ確認する文字数をずらす
            sp_str = Left(sp_str, idx_chr - 1) & addDq(count_dq / 2) & Right(sp_str, Len(sp_str) - idx_chr - count_dq + 1)
            idx_chr = idx_chr + (count_dq / 2)
        Else
            If str_chr = "," Then
                If flg_str = False Then
                    word(UBound(word)) = Mid(sp_str, pos_start, idx_chr - pos_start)
                    If word(UBound(word)) = """" Then
                        'カンマで区切られた内容として、""だった場合は、空文字
                        word(UBound(word)) = ""
                    End If
                    ReDim Preserve word(1 To UBound(word) + 1)
                    pos_start = idx_chr + 1
                End If
            End If
            idx_chr = idx_chr + 1
        End If
        
        If idx_chr > Len(sp_str) Then
            word(UBound(word)) = Mid(sp_str, pos_start)
            If flg_str = False Then
                If word(UBound(word)) = """" Then
                    'カンマで区切られた内容として、""だった場合は、空文字
                    word(UBound(word)) = ""
                End If
            End If
        End If
    Loop
    
    For idx_wd = LBound(word) To UBound(word)
        word(idx_wd) = Replace(word(idx_wd), "\\", "\")
    Next idx_wd
    
    splitCsv = word
    
End Function

Function doubleQuatCount(ByVal tar_str As String, ByVal idx As Long) As Long
    '連続するダブルクォーテーションの数を取得する
    Dim dq_count As Long
    Dim idx_chr As Long
    
    dq_count = 0
    For idx_chr = idx To Len(tar_str)
        If Mid(tar_str, idx_chr, 1) = """" Then
            dq_count = dq_count + 1
        Else
            Exit For
        End If
    Next idx_chr
    doubleQuatCount = dq_count
End Function

Function addDq(ByVal count As Long) As String
    Dim dq As String
    Dim idx_count As Long
    
    dq = ""
    For idx_count = 1 To count
        dq = dq & """"
    Next idx_count
    addDq = dq
End Function