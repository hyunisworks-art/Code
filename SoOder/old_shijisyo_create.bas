
'このモジュールは開発の過程で一時的にバックアップとして旧モジュールコードを保存したものです。機能はすべて移動されているため、将来的に不要になる可能性があります。（2026-01 西村記載）

Sub create_shijisyo()
    
    '「作業指示書作成リスト」シート保護解除
    ThisWorkbook.Worksheets(shijisyo_list_sheet).Unprotect
    
    'パターン分けの色を指定
    Dim patern_1_color As Long
    patern_1_color = RGB(color1_R, color1_G, color1_B) '紫色
    Dim patern_2_color As Long
    patern_2_color = RGB(color2_R, color2_G, color2_B) '緑色
    Dim patern_3_color As Long
    patern_3_color = RGB(color3_R, color3_G, color3_B) '赤色
    Dim patern_4_color As Long
    patern_4_color = RGB(color4_R, color4_G, color4_B) '黄色
    
    Dim my_dir As String
    Dim my_file As String
    my_file = ThisWorkbook.path & "\" & ThisWorkbook.Name
    Dim rng_thisworkbook As Range
    Set rng_thisworkbook = ThisWorkbook.Worksheets(file_name_save_sheet).Cells(2, 1)
    my_file = GetFileName(my_file, rng_thisworkbook)
    my_dir = Left(my_file, InStrRev(my_file, "\"))
    
    With ThisWorkbook.Worksheets(file_name_save_sheet)
        Dim csv_path As String
        csv_path = .Cells(1, 1).Value
    End With
    
    'CSVファイルの存在確認（無い場合はメッセージ表示して終了）
    '「w2pデータ取り込み」ボタンで指定されたファイルアドレスを「ファイル名保持」シートにあらかじめ保存している
    If csv_path = "" Then
        MsgBox "w2pデータが取り込めていません。" & vbCrLf & "先に「w2pデータ取り込み」ボタンを押下してください。"
        End
    ElseIf Dir(csv_path, vbNormal) = "" Then
        MsgBox "「w2pデータ取り込み」で指定されたcsvファイルが存在しません。" & vbCrLf & "先に「w2pデータ取り込み」ボタンを押下してください。"
        End
    Else
        '指定したCSVファイルを読み込む（UTF-8形式対応）
        Dim csv_data_obj As csv_data
        Dim csv_sep As Object
        Set csv_sep = CreateObject("ADODB.stream")
            csv_sep.charset = "utf-8"
            csv_sep.Open
            csv_sep.LoadFromFile (csv_path)
        Dim sp_str As String
        sp_str = csv_sep.ReadText
        Dim sp_str_line As Variant
        sp_str_line = Split(sp_str, vbCrLf)
        
        ReDim csv_data_obj.csv_list(1 To UBound(sp_str_line) + 1, 1 To 1)
        
        Dim line_sep_tab() As Variant
        
        '読み込んだCSVファイルを2次元配列（csv_data_obj.csv_list）に格納
        For now_row = 0 To UBound(sp_str_line)
            If InStr(sp_str_line(now_row), ",") <> 0 Then
                line_sep_tab = splitCsv(sp_str_line(now_row))
                If now_row = 0 Then
                    ReDim Preserve csv_data_obj.csv_list(1 To UBound(sp_str_line) + 1, 1 To UBound(line_sep_tab))
                End If
                For now_clm = 1 To UBound(line_sep_tab)
                    csv_data_obj.csv_list(now_row + 1, now_clm) = line_sep_tab(now_clm)
                Next
            End If
        Next
        
        'CSVファイルの列番号設定（CSV列数に応じて W2P / Spinno のマッピングを切り替える）
        csv_data_obj.title_row = 1
        Dim rawCsv As Variant
        rawCsv = csv_data_obj.csv_list
        Dim rawCsvColCount As Long
        rawCsvColCount = UBound(rawCsv, 2)
        Dim csvColCount As Long
        csvColCount = rawCsvColCount

        '列数によりマッピングを切り替え（W2P=143列 / Spinno=20列なので、100列を目安に判断する）
        If rawCsvColCount > 100 Then
            ' W2P形式の列マッピング
            csv_data_obj.store_clm = 1
            csv_data_obj.order_nom_clm = 2
            csv_data_obj.order_status_clm = 19
            csv_data_obj.item_code_clm = 21
            csv_data_obj.item_name_clm = 22
            csv_data_obj.item_count_clm = 23
            csv_data_obj.haisou_order_name_clm = 8
            csv_data_obj.haisou_name_clm = 12
            csv_data_obj.haisou_post_clm = 13
            csv_data_obj.haisou_address1_clm = 14
            csv_data_obj.haisou_address2_clm = 15
            csv_data_obj.haisou_address3_clm = 16
            csv_data_obj.haisou_tel_clm = 17
            csv_data_obj.haisousaki_tantou_clm = 18
        Else
            'Spinno形式の列マッピング
            csv_data_obj.store_clm = 1
            csv_data_obj.order_nom_clm = 2
            csv_data_obj.order_date_clm = 4
            csv_data_obj.haisou_order_name_clm = 8
            csv_data_obj.haisou_name_clm = 11
            csv_data_obj.haisou_post_clm = 12
            csv_data_obj.haisou_address1_clm = 13
            csv_data_obj.haisou_address2_clm = 14
            csv_data_obj.haisou_address3_clm = 15
            csv_data_obj.haisou_tel_clm = 16
            csv_data_obj.order_status_clm = 17
            csv_data_obj.item_code_clm = 18
            csv_data_obj.item_name_clm = 19
            csv_data_obj.item_count_clm = 20
            csv_data_obj.haisousaki_tantou_clm = 0 '該当なし
        End If
        
        '「W2Pデータ貼り付け」シートの列番号設定
        Dim w2p_data As w2p_data
        w2p_data.title_row = 1
        w2p_data.store_clm = 1
        w2p_data.order_nom_clm = 2
        w2p_data.order_date_clm = 4
        w2p_data.haisou_order_name_clm = 8
        w2p_data.order_status_clm = 19
        w2p_data.item_code_clm = 20
        w2p_data.item_name_clm = 21
        w2p_data.item_count_clm = 22
        w2p_data.haisou_name_clm = 12
        w2p_data.haisou_post_clm = 13
        w2p_data.haisou_address1_clm = 14
        w2p_data.haisou_address2_clm = 15
        w2p_data.haisou_address3_clm = 16
        w2p_data.haisou_tel_clm = 17
        w2p_data.haisousaki_tantou_clm = 18
        w2p_data.haisou_name_item_clm = 25
        w2p_data.haisou_tantou_item_clm = 31
        w2p_data.nouki_clm = 32
        w2p_data.sagyou_shiji_clm = 33
        w2p_data.syukko_yotei_clm = 34
        ' w2p_data.end_data_clm will be determined dynamically from the W2P sheet header
        w2p_data.end_data_clm = 0
        
        '「ファイル名設定」シートの内容取得
        Dim set_file_name_data As set_file_name_data
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
        
        '「w2pデータ貼り付け」シートのデータ取得
        Dim actualEndCol As Long
        With ThisWorkbook.Worksheets(w2pdata_sheet)
            ' determine how many columns the W2P sheet exposes (header row width)
            actualEndCol = .Cells(1, .Columns.count).End(xlToLeft).Column
            If actualEndCol > 0 Then
                w2p_data.end_data_clm = actualEndCol
            Else
                ' fallback to previous default if something goes wrong
                If w2p_data.end_data_clm <= 0 Then w2p_data.end_data_clm = 39
            End If
            ReDim w2p_data.w2p_list(1 To UBound(csv_data_obj.csv_list, 1), 1 To w2p_data.end_data_clm)
            w2p_data.w2p_list = .Range(.Cells(1, 1), .Cells(UBound(w2p_data.w2p_list, 1), w2p_data.end_data_clm)).Value
        End With

    'Spinno形式のCSVデータをW2P形式にマッピングする処理
    Dim csvWasMapped As Boolean
    csvWasMapped = False

    'CSVの列数が100未満の場合のみ、この処理を行う（Spinno形式）
    If rawCsvColCount < 100 Then
            'CSVデータを別構造にマッピングするための配列
            Dim mappedCsv() As Variant
            '行数・列数の最大値用変数
            Dim rMax As Long, cMax As Long
            'rawCsvの行数取得用（エラー対策込み）
            Dim rMaxCandidate As Long
            'rawCsvが未初期化などでもエラーで止まらないようにする
            On Error Resume Next
            'rawCsvの1次元目（行数）の最大インデックスを取得
            rMaxCandidate = UBound(rawCsv, 1)
            'エラーが出た場合の処理（rawCsvが空など）
            If Err.Number <> 0 Then
                Err.Clear
                rMaxCandidate = 0
            End If
            On Error GoTo 0
            If rMaxCandidate < 1 Then
                rMax = 1
            Else
                rMax = rMaxCandidate
            End If

            'ヘッダー行の期待値を定義（W2P形式のヘッダー：Spinno形式は列数が少ないためW2P形式にマッピング）
            Dim expectedHeaders As Variant
            Dim expectedHeadersCsv As String
            'ヘッダー行の期待値文字列（区切り文字は「|」）
            expectedHeadersCsv = "ストア|発注番号|明細番号|注文日|顧客|グループ|発注者ID|発注者|発注者ログインID|発注者コード|件名|配送先名|配送先郵便番号|配送先住所1|配送先住所2|配送先住所3|配送先電話番号|配送先担当者名|" & _
                                 "注文状態|本発注日時|商品コード|商品名|注文数量|単価|小計|注文小計|消費税|合計（税込）|配送先名(ｱｲﾃﾑ別）|配送先郵便番号(ｱｲﾃﾑ別）|配送先住所1(ｱｲﾃﾑ別）|配送先住所2(ｱｲﾃﾑ別）|配送先住所3(ｱｲﾃﾑ別）|配送先電話番号(ｱｲﾃﾑ別）|配送先担当者名(ｱｲﾃﾑ別）|入稿データ1|入稿データ2|入稿データ3|" & _
                                 "入稿データ4|入稿データ5|入稿データ6|入稿データ7|入稿データ8|入稿データ9|入稿データ10|配送業者番号|出荷日|配送備考|対象ストア|略称|JANコード|説明|備考|納期テキスト|納期（標準）|納期（最短）|納期（締め切り時間）|サムネイル|アイテム詳細画像|納品タイプ|印刷版データ|" & _
                                 "自由項目名1|自由項目値1|自由項目名2|自由項目値2|自由項目名3|自由項目値3|自由項目名4|自由項目値4|自由項目名5|自由項目値5|自由項目名6|自由項目値6|自由項目名7|自由項目値7|自由項目名8|自由項目値8|自由項目名9|自由項目値9|自由項目名10|自由項目値10|公開|プロダクトコード|プロダクト名|プロダクト分類|説明（商品プロダクト）|デフォルトプロダクション|" & _
                                 "パーツ名1|版データ1|パーツ名2|版データ2|パーツ名3|版データ3|パーツ名4|版データ4|パーツ名5|版データ5|仕上がりサイズ|パーツ名1|裁ち1|色1表|色1裏|用紙1|パーツ名2|裁ち2|色2表|色2裏|用紙2|パーツ名3|裁ち3|色3表|色3裏|用紙3|" & _
                                 "パーツ名4|裁ち4|色4表|色4裏|用紙4|パーツ名5|裁ち5|色5表|色5裏|用紙5|折り|特記事項（折）|製本|製本部材|製本部材カラー|特記事項（製本）|穴あけ|特記事項（穴あけ）|断裁|特記事項（断裁)|" & _
                                 "その他加工1|特記事項（その他加工1)|その他加工2|特記事項（その他加工2)|その他加工3|特記事項（その他加工3)|その他加工4|特記事項（その他加工4)|その他加工5|特記事項（その他加工5)"
            '期待するCSVヘッダー文字列を「|」区切りで配列に変換
            expectedHeaders = Split(expectedHeadersCsv, "|")

            'expectedHeaders の要素数を安全に取得（未初期化・空配列でもエラーで止まらないように）
            Dim expCount As Long
            expCount = 0
            '配列の要素数 = 上限 - 下限 + 1
            On Error Resume Next
            expCount = UBound(expectedHeaders) - LBound(expectedHeaders) + 1
            If Err.Number <> 0 Then
                Err.Clear
                expCount = 0
            End If
            On Error GoTo 0
            'W2Pデータ（2次元配列）の列数を安全に取得
            Dim w2pCols As Long
            On Error Resume Next
            w2pCols = UBound(w2p_data.w2p_list, 2)
            If Err.Number <> 0 Then
                Err.Clear
                w2pCols = 0
            End If
            On Error GoTo 0
        
            'mappedCsv の列数は expectedHeaders の要素数に固定する
            cMax = 0
            ' expectedHeaders の要素数を安全に取得
            On Error Resume Next
            cMax = UBound(expectedHeaders) - LBound(expectedHeaders) + 1
            If Err.Number <> 0 Then
                Err.Clear
                cMax = 0
            End If
            On Error GoTo 0
            ' 念のため最低1列は確保
            If cMax < 1 Then cMax = 1
            ' 行数 rMax × 列数 cMax の2次元配列を確保
            ReDim mappedCsv(1 To rMax, 1 To cMax)
            '1) ヘッダー行を作成（expectedHeaders をそのまま1行目にセット）
            For hc = 1 To cMax
                mappedCsv(1, hc) = CStr(expectedHeaders(hc - 1))
            Next hc

            '2) データ行を rawCsv → mappedCsv に列マッピングしてコピー
            Dim rr As Long
            For rr = 2 To rMax
                'ストアID列：固定文字列を付与して格納
                If csv_data_obj.store_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.store_clm) = "SOMPOPA@" & rawCsv(rr, csv_data_obj.store_clm)
                '発注番号列：そのままコピー
                If csv_data_obj.order_nom_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.order_nom_clm) = rawCsv(rr, csv_data_obj.order_nom_clm)
                ' 注文内部管理番号（rawCsvの3列目）→ mappedCsvの3列目（明細番号扱い）へコピー
                If 3 <= UBound(mappedCsv, 2) And 3 <= UBound(rawCsv, 2) Then
                    mappedCsv(rr, 3) = rawCsv(rr, 3)
                End If
                ' 発注依頼日（注文日）→ mappedCsv の注文日列へそのままコピー
                If csv_data_obj.order_date_clm <= UBound(rawCsv, 2) _
                    And w2p_data.order_date_clm <= UBound(mappedCsv, 2) Then
                    mappedCsv(rr, w2p_data.order_date_clm) = CStr(rawCsv(rr, csv_data_obj.order_date_clm))
                End If
                ' 顧客会社名 → 5列目（固定列）にコピー
                If 5 <= UBound(mappedCsv, 2) Then
                    Dim customer_company_clm As Long
                    customer_company_clm = 5
                    ' rawCsv の同列(5列目)から mappedCsv の5列目へコピー
                    If customer_company_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, 5) = rawCsv(rr, customer_company_clm)
                End If
                'グループ名 → 6列目（固定列）にコピー
                If 6 <= UBound(mappedCsv, 2) Then
                    Dim group_name_clm As Long
                    group_name_clm = 6
                    'If group_name_clm <= UBound(csv_data_obj.csv_list, 2) Then mappedCsv(rr, 6) = csv_data_obj.csv_list(rr, group_name_clm)
                    If group_name_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, 6) = rawCsv(rr, group_name_clm)
                End If
                ' ユーザー番号 → 7列目（固定列）にコピー
                If 7 <= UBound(mappedCsv, 2) Then
                    Dim user_number_clm As Long
                    user_number_clm = 7
                    If user_number_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, 7) = rawCsv(rr, user_number_clm)
                End If
                ' 発注者名 → w2p_data.haisou_order_name_clm（W2P側の列定義）へコピー
                If csv_data_obj.haisou_order_name_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.haisou_order_name_clm) = rawCsv(rr, csv_data_obj.haisou_order_name_clm)
                '件名列：そのままコピー
                If 9 <= UBound(mappedCsv, 2) And 9 <= UBound(rawCsv, 2) Then mappedCsv(rr, 9) = rawCsv(rr, 9)
                If 10 <= UBound(mappedCsv, 2) And 10 <= UBound(rawCsv, 2) Then mappedCsv(rr, 10) = rawCsv(rr, 10)
                '配送先担当者名列：そのままコピー
                If csv_data_obj.haisou_name_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.haisou_name_clm) = rawCsv(rr, csv_data_obj.haisou_name_clm)
                '配送先郵便番号列：そのままコピー
                If csv_data_obj.haisou_post_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.haisou_post_clm) = rawCsv(rr, csv_data_obj.haisou_post_clm)
                '配送先住所1列：そのままコピー
                If csv_data_obj.haisou_address1_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.haisou_address1_clm) = rawCsv(rr, csv_data_obj.haisou_address1_clm)
                '配送先住所2列：そのままコピー
                If csv_data_obj.haisou_address2_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.haisou_address2_clm) = rawCsv(rr, csv_data_obj.haisou_address2_clm)
                '配送先住所3列：コピー時に「？」を「-」に置換してから格納
                If csv_data_obj.haisou_address3_clm <= UBound(rawCsv, 2) Then
                Dim tmpAddr3 As String
                    tmpAddr3 = CStr(rawCsv(rr, csv_data_obj.haisou_address3_clm))
                    tmpAddr3 = Replace(tmpAddr3, "？", "-")
                    tmpAddr3 = Replace(tmpAddr3, "?", "-")
                    mappedCsv(rr, w2p_data.haisou_address3_clm) = tmpAddr3
                End If
                '配送先住所2・3列の結合処理
                On Error Resume Next
                If w2p_data.haisou_address2_clm >= 1 And w2p_data.haisou_address3_clm >= 1 Then
                    Dim a2 As String, a3 As String
                    a2 = CStr(mappedCsv(rr, w2p_data.haisou_address2_clm))
                    a3 = CStr(mappedCsv(rr, w2p_data.haisou_address3_clm))
                    If Trim(a3) <> "" Then
                        mappedCsv(rr, w2p_data.haisou_address2_clm) = Trim(a2 & " " & a3)
                        mappedCsv(rr, w2p_data.haisou_address3_clm) = ""
                    End If
                End If
                On Error GoTo 0
                '配送先電話番号列：そのままコピー
                If csv_data_obj.haisou_tel_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.haisou_tel_clm) = rawCsv(rr, csv_data_obj.haisou_tel_clm)
                '注文状態列：そのままコピー
                If csv_data_obj.order_status_clm <= UBound(rawCsv, 2) Then mappedCsv(rr, w2p_data.order_status_clm) = rawCsv(rr, csv_data_obj.order_status_clm)
                '商品コード列:固定21列目にコピー
                If csv_data_obj.item_code_clm <= UBound(rawCsv, 2) And 21 <= UBound(mappedCsv, 2) Then
                    mappedCsv(rr, 21) = rawCsv(rr, csv_data_obj.item_code_clm)
                End If
                ' 商品名 → 固定で22列目にコピー
                If csv_data_obj.item_name_clm <= UBound(rawCsv, 2) And 22 <= UBound(mappedCsv, 2) Then
                    mappedCsv(rr, 22) = rawCsv(rr, csv_data_obj.item_name_clm)
                End If
                ' 数量 → 固定で23列目にコピー
                If csv_data_obj.item_count_clm <= UBound(rawCsv, 2) And 23 <= UBound(mappedCsv, 2) Then
                    mappedCsv(rr, 23) = rawCsv(rr, csv_data_obj.item_count_clm)
                End If
            Next rr

            'マッピング済み配列を csv_data_obj の実体として差し替え
            csv_data_obj.csv_list = mappedCsv
            csvWasMapped = True
            csvColCount = cMax
            ' Update csv_data column indices to W2P positions so later logic uses correct columns
            csv_data_obj.store_clm = w2p_data.store_clm
            csv_data_obj.order_nom_clm = w2p_data.order_nom_clm
            csv_data_obj.order_date_clm = w2p_data.order_date_clm
            csv_data_obj.haisou_order_name_clm = w2p_data.haisou_order_name_clm
            csv_data_obj.haisou_name_clm = w2p_data.haisou_name_clm
            csv_data_obj.haisou_post_clm = w2p_data.haisou_post_clm
            csv_data_obj.haisou_address1_clm = w2p_data.haisou_address1_clm
            csv_data_obj.haisou_address2_clm = w2p_data.haisou_address2_clm
            csv_data_obj.haisou_address3_clm = w2p_data.haisou_address3_clm
            csv_data_obj.haisou_tel_clm = w2p_data.haisou_tel_clm
            csv_data_obj.order_status_clm = w2p_data.order_status_clm
            csv_data_obj.item_code_clm = w2p_data.item_code_clm
            csv_data_obj.item_name_clm = w2p_data.item_name_clm
            csv_data_obj.item_count_clm = w2p_data.item_count_clm
            csv_data_obj.haisousaki_tantou_clm = w2p_data.haisousaki_tantou_clm
            Else
            csv_data_obj.csv_list = rawCsv
            csvColCount = rawCsvColCount
        End If
        
        '「定款コード」シートの内容取得
        Dim t_code_data As teikan_code_data
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
        
        '「作業指示書作成リスト」シートにデータを格納
        With ThisWorkbook.Worksheets(shijisyo_list_sheet)
            Dim shijisyo_list_data As shijisyo_list_data
            shijisyo_list_data.shijisyo_title_row = 1
            With .Cells(1, 1).SpecialCells(xlLastCell)
                shijisyo_list_data.end_row = .row
                shijisyo_list_data.end_clm = .Column
            End With
            ReDim shijisyo_list_data.shijisyo_list(1 To UBound(w2p_data.w2p_list, 1), 1 To shijisyo_list_data.end_clm)
            ' 配列の列数が書き込みで参照する列数より小さいとエラーになるため拡張
            Dim requiredCols As Long
            requiredCols = shijisyo_list_data.end_clm
            ' 参照予定の列番号（現在の設定から）を確認して最大値を求める
            If 4 > requiredCols Then requiredCols = 4
            If 5 > requiredCols Then requiredCols = 5
            If 8 > requiredCols Then requiredCols = 8
            If 9 > requiredCols Then requiredCols = 9
            If 10 > requiredCols Then requiredCols = 10
            If 11 > requiredCols Then requiredCols = 11
            If 12 > requiredCols Then requiredCols = 12
            If 13 > requiredCols Then requiredCols = 13
            If 14 > requiredCols Then requiredCols = 14
            If 15 > requiredCols Then requiredCols = 15
            If 17 > requiredCols Then requiredCols = 17
            If 20 > requiredCols Then requiredCols = 20
            If 26 > requiredCols Then requiredCols = 26
            If 27 > requiredCols Then requiredCols = 27
            If 28 > requiredCols Then requiredCols = 28
            If requiredCols > UBound(shijisyo_list_data.shijisyo_list, 2) Then
                ReDim Preserve shijisyo_list_data.shijisyo_list(1 To UBound(shijisyo_list_data.shijisyo_list, 1), 1 To requiredCols)
            End If
            'タイトル取得
            For now_clm = 1 To UBound(shijisyo_list_data.shijisyo_list, 2)
                shijisyo_list_data.shijisyo_list(shijisyo_list_data.shijisyo_title_row, now_clm) = .Cells(shijisyo_list_data.shijisyo_title_row, now_clm).Value
            Next
            'シート初期化
            If shijisyo_list_data.end_row > 1 Then
                .Rows("2:" & shijisyo_list_data.end_row).Clear
            End If
            
            '「作業指示書作成リスト」シートの列番号設定
            shijisyo_list_data.item_code_clm = 4
            shijisyo_list_data.item_name_clm = 5
            shijisyo_list_data.order_name_clm = 8
            shijisyo_list_data.haisou_name_clm = 9
            shijisyo_list_data.haisou_tanou_clm = 10
            shijisyo_list_data.haisou_tel_clm = 11
            shijisyo_list_data.haisou_post_clm = 12
            shijisyo_list_data.haisou_address1_clm = 13
            shijisyo_list_data.haisou_address2_clm = 14
            shijisyo_list_data.haisou_address3_clm = 15
            shijisyo_list_data.item_count_clm = 17
            shijisyo_list_data.order_nom_clm = 20
            shijisyo_list_data.nouki_clm = 26
            shijisyo_list_data.sagyou_shiji_clm = 27
            shijisyo_list_data.syukko_yotei_clm = 28
            
            Dim shijisyo_count As Long
            shijisyo_count = 1
            '商品コードが数値のみのものを切り出し
            For now_csv_row = 2 To UBound(csv_data_obj.csv_list, 1)
                'CSVに item_code 列が存在するか確認（Spinno/W2Pで列位置が異なるため）
                If csv_data_obj.item_code_clm <= UBound(csv_data_obj.csv_list, 2) Then
                    If IsNumeric(csv_data_obj.csv_list(now_csv_row, csv_data_obj.item_code_clm)) = True Then
                        If IsEmpty(csv_data_obj.csv_list(now_csv_row, csv_data_obj.item_code_clm)) = False Then
                            '納期が空の場合=「定款」の場合、「作業指示書作成リスト」シートには記入しない
                            If w2p_data.w2p_list(now_csv_row, w2p_data.nouki_clm) <> "" Then
                            shijisyo_count = shijisyo_count + 1
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.order_nom_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.order_nom_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.order_name_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_order_name_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_name_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_name_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_post_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_post_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_address1_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_address1_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_address2_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_address2_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_address3_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_address3_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_tel_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisou_post_clm)
                            If csv_data_obj.haisousaki_tantou_clm >= 1 And csv_data_obj.haisousaki_tantou_clm <= UBound(csv_data_obj.csv_list, 2) Then
                                shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_tanou_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.haisousaki_tantou_clm)
                            Else
                                shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.haisou_tanou_clm) = ""
                            End If
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.item_code_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.item_code_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.item_name_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.item_name_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.item_count_clm) = csv_data_obj.csv_list(now_csv_row, csv_data_obj.item_count_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.nouki_clm) = w2p_data.w2p_list(now_csv_row, w2p_data.nouki_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.sagyou_shiji_clm) = w2p_data.w2p_list(now_csv_row, w2p_data.sagyou_shiji_clm)
                            shijisyo_list_data.shijisyo_list(shijisyo_count, shijisyo_list_data.syukko_yotei_clm) = w2p_data.w2p_list(now_csv_row, w2p_data.syukko_yotei_clm)
                        End If
                    End If
                End If
                End If
            Next
            .Range(.Cells(1, 1), .Cells(UBound(shijisyo_list_data.shijisyo_list, 1), UBound(shijisyo_list_data.shijisyo_list, 2))) = shijisyo_list_data.shijisyo_list
        End With
        
        'パターン1に該当するデータが存在する場合のみ、作業指示書と注文書作成
        If shijisyo_count >= 2 Then
            'テンプレートファイルが存在した場合、作業指示書作成機能を呼び出す
            If Dir(my_dir & "\" & temp_file) <> "" Then
                Call Create_Shijisho
            Else
                MsgBox temp_file & "が見つかりません。" & vbCrLf & "作業指示書の作成に失敗しました。"
            End If
            
            '注文書作成
            With ThisWorkbook
                Dim f_name As String
                '注文書ファイル名設定
                f_name = set_file_name_data.file_name_list(set_file_name_data.order_list_row, set_file_name_data.file_name_clm)
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
        End If
        
        'csv出力
        With ThisWorkbook.Worksheets(w2pdata_sheet)
            Dim sep_w2p_data As sep_w2p_data
            sep_w2p_data.shindou_list_count = 1
            sep_w2p_data.maru_kyoten_list_count = 1
            sep_w2p_data.maru_list_count = 1
            sep_w2p_data.teikan_list_count = 1
            sep_w2p_data.shindou_list_title_row = 1
            sep_w2p_data.maru_kyoten_list_title_row = 1
            sep_w2p_data.maru_list_title_row = 1
            sep_w2p_data.teikan_list_title_row = 1
            
            ReDim sep_w2p_data.shindou_list(1 To UBound(w2p_data.w2p_list, 1), 1 To UBound(w2p_data.w2p_list, 2))
            ReDim sep_w2p_data.maru_kyoten_list(1 To UBound(w2p_data.w2p_list, 1), 1 To UBound(w2p_data.w2p_list, 2))

            '中間ファイル用配列初期化
            Dim middleCols As Long
            Dim middleHeaders() As String
            middleCols = UBound(csv_data_obj.csv_list, 2)
            If middleCols < 100 Then
                Dim midFolder As String, midFile As String
                midFolder = Left(csv_path, InStrRev(csv_path, "\"))
                midFile = Dir(midFolder & "S*中間ファイル.csv")
                If midFile <> "" Then
                    Dim stm As Object, hdrLine As String
                    Set stm = CreateObject("ADODB.stream")
                    stm.charset = "utf-8"
                    stm.Open
                    stm.LoadFromFile midFolder & midFile
                    hdrLine = stm.ReadText(adReadLine)
                    stm.Close
                    Set stm = Nothing
                    If Len(Trim(hdrLine)) > 0 Then
                        middleHeaders = Split(hdrLine, ",")
                        middleCols = UBound(middleHeaders) + 1
                    End If
                Else
                    ' try workbook folder as fallback (project repository)
                    Dim wbFolder As String, wbFile As String
                    wbFolder = ThisWorkbook.path & "\"
                    wbFile = Dir(wbFolder & "S*中間ファイル.csv")
                    If wbFile <> "" Then
                        Dim stm2 As Object, hdrLine2 As String
                        Set stm2 = CreateObject("ADODB.stream")
                        stm2.charset = "utf-8"
                        stm2.Open
                        stm2.LoadFromFile wbFolder & wbFile
                        hdrLine2 = stm2.ReadText(adReadLine)
                        stm2.Close
                        Set stm2 = Nothing
                        If Len(Trim(hdrLine2)) > 0 Then
                            middleHeaders = Split(hdrLine2, ",")
                            middleCols = UBound(middleHeaders) + 1
                        End If
                    End If
                End If
            End If
            ReDim sep_w2p_data.maru_middle_list(1 To UBound(w2p_data.w2p_list, 1), 1 To middleCols)

            '中間ファイル用データマッピング
            Dim headerMap() As Long
            Dim headerMapAllocated As Boolean
            headerMapAllocated = False
            If middleCols > 0 Then
                ReDim headerMap(1 To middleCols)
                headerMapAllocated = True
                Dim iHdr As Long, jCsv As Long
                If IsArray(middleHeaders) Then
                    Dim mhL As Long, mhU As Long
                    On Error Resume Next
                    mhL = LBound(middleHeaders)
                    If Err.Number <> 0 Then
                        ' array not allocated or invalid; clear error and mark as empty
                        Err.Clear
                        On Error GoTo 0
                        mhL = 1: mhU = 0
                    Else
                        mhU = UBound(middleHeaders)
                        On Error GoTo 0
                    End If
                    ' only iterate if array has at least one element
                    If mhU >= mhL Then
                        For iHdr = 1 To UBound(headerMap)
                            headerMap(iHdr) = 0
                            Dim tmplName As String
                            Dim idx As Long
                            idx = iHdr - 1 + mhL
                            If idx >= mhL And idx <= mhU Then
                                tmplName = Replace(middleHeaders(idx), Chr(34), "")
                            Else
                                tmplName = ""
                            End If
                            ' normalize template name
                            tmplName = Trim(tmplName)
                            tmplName = LCase(tmplName)
                            tmplName = Replace(tmplName, "　", " ")
                            tmplName = Replace(tmplName, "（", "(")
                            tmplName = Replace(tmplName, "）", ")")
                            If tmplName <> "" Then
                                Dim csvHdrNorm As String
                                ' 1) exact normalized match
                                For jCsv = 1 To UBound(csv_data_obj.csv_list, 2)
                                    csvHdrNorm = LCase(Trim(CStr(csv_data_obj.csv_list(csv_data_obj.title_row, jCsv))))
                                    csvHdrNorm = Replace(csvHdrNorm, "　", " ")
                                    csvHdrNorm = Replace(csvHdrNorm, "（", "(")
                                    csvHdrNorm = Replace(csvHdrNorm, "）", ")")
                                    If csvHdrNorm = tmplName Then
                                        headerMap(iHdr) = jCsv
                                        Exit For
                                    End If
                                Next jCsv
                                ' 2) partial match fallback
                                If headerMap(iHdr) = 0 Then
                                    For jCsv = 1 To UBound(csv_data_obj.csv_list, 2)
                                        csvHdrNorm = LCase(Trim(CStr(csv_data_obj.csv_list(csv_data_obj.title_row, jCsv))))
                                        csvHdrNorm = Replace(csvHdrNorm, "　", " ")
                                        If InStr(csvHdrNorm, tmplName) > 0 Or InStr(tmplName, csvHdrNorm) > 0 Then
                                            headerMap(iHdr) = jCsv
                                            Exit For
                                        End If
                                    Next jCsv
                                End If
                                ' 3) If still not mapped and this is Spinno (no proper csv header), try keyword -> known csv_data column mapping
                                If headerMap(iHdr) = 0 Then
                                    'If csvColCount < 100 Then
                                    If rawCsvColCount < 100 Then
                                        Dim tKey As String
                                        tKey = tmplName
                                        ' simple keyword based mapping to Spinno known column indices
                                        If InStr(tKey, "発注番号") > 0 Then headerMap(iHdr) = csv_data_obj.order_nom_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "店舗") > 0 Or InStr(tKey, "ストア") > 0) Then headerMap(iHdr) = csv_data_obj.store_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "商品コード") > 0 Or InStr(tKey, "品番") > 0 Or InStr(tKey, "商品ｺｰﾄﾞ") > 0) Then headerMap(iHdr) = csv_data_obj.item_code_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "商品名") > 0 Or InStr(tKey, "品名") > 0) Then headerMap(iHdr) = csv_data_obj.item_name_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "数量") > 0 Or InStr(tKey, "個数") > 0) Then headerMap(iHdr) = csv_data_obj.item_count_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "郵便") > 0 Or InStr(tKey, "郵便番号") > 0) Then headerMap(iHdr) = csv_data_obj.haisou_post_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "住所1") > 0 Or InStr(tKey, "住所１") > 0 Or InStr(tKey, "都道府県") > 0) Then headerMap(iHdr) = csv_data_obj.haisou_address1_clm
                                        If headerMap(iHdr) = 0 And InStr(tKey, "住所2") > 0 Then headerMap(iHdr) = csv_data_obj.haisou_address2_clm
                                        If headerMap(iHdr) = 0 And InStr(tKey, "住所3") > 0 Then headerMap(iHdr) = csv_data_obj.haisou_address3_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "電話") > 0 Or InStr(tKey, "tel") > 0) Then headerMap(iHdr) = csv_data_obj.haisou_tel_clm
                                        If headerMap(iHdr) = 0 And (InStr(tKey, "宛名") > 0 Or InStr(tKey, "配送先") > 0 Or InStr(tKey, "お届け") > 0) Then headerMap(iHdr) = csv_data_obj.haisou_name_clm
                                        If headerMap(iHdr) = 0 And InStr(tKey, "納期") > 0 Then headerMap(iHdr) = w2p_data.nouki_clm
                                        If headerMap(iHdr) = 0 And InStr(tKey, "注文日") > 0 Then headerMap(iHdr) = csv_data_obj.order_date_clm
                                    End If
                                End If
                            End If
                        Next iHdr
                    End If
                    ' if no mapping found, fallback to sequential mapping where possible
                    Dim anyMapped As Boolean
                    anyMapped = False
                    For iHdr = 1 To UBound(headerMap)
                        If headerMap(iHdr) <> 0 Then anyMapped = True: Exit For
                    Next iHdr
                    If Not anyMapped Then
                        Dim maxMap As Long
                        If UBound(headerMap) < UBound(csv_data_obj.csv_list, 2) Then
                            maxMap = UBound(headerMap)
                        Else
                            maxMap = UBound(csv_data_obj.csv_list, 2)
                        End If
                        For iHdr = 1 To maxMap
                            headerMap(iHdr) = iHdr
                        Next iHdr
                    End If
                Else
                    ' no template headers; fallback to sequential mapping up to available columns
                    Dim maxMap2 As Long
                    If middleCols < UBound(csv_data_obj.csv_list, 2) Then
                        maxMap2 = middleCols
                    Else
                        maxMap2 = UBound(csv_data_obj.csv_list, 2)
                    End If
                    For iHdr = 1 To maxMap2
                        headerMap(iHdr) = iHdr
                    Next iHdr
                End If
            End If
            '中間ファイル用データマッピングおわり

            'タイトル行挿入用変数
            ReDim sep_w2p_data.maru_list(1 To UBound(w2p_data.w2p_list, 1), 1 To UBound(w2p_data.w2p_list, 2))
            ReDim sep_w2p_data.teikan_list(1 To UBound(w2p_data.w2p_list, 1), 1 To UBound(w2p_data.w2p_list, 2))
            
            'タイトル行挿入
            For now_w2p_title_clm = 1 To UBound(w2p_data.w2p_list, 2)
                sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_title_row, now_w2p_title_clm) = w2p_data.w2p_list(w2p_data.title_row, now_w2p_title_clm)
                sep_w2p_data.maru_kyoten_list(sep_w2p_data.shindou_list_title_row, now_w2p_title_clm) = w2p_data.w2p_list(w2p_data.title_row, now_w2p_title_clm)
                sep_w2p_data.maru_list(sep_w2p_data.maru_list_title_row, now_w2p_title_clm) = w2p_data.w2p_list(w2p_data.title_row, now_w2p_title_clm)
                sep_w2p_data.teikan_list(sep_w2p_data.teikan_list_title_row, now_w2p_title_clm) = w2p_data.w2p_list(w2p_data.title_row, now_w2p_title_clm)
            Next
            '中間ファイルのタイトル取得
            If middleCols >= 100 Or UBound(csv_data_obj.csv_list, 2) >= 100 Then
                For now_orderdetail_title_clm = 1 To UBound(csv_data_obj.csv_list, 2)
                    sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, now_orderdetail_title_clm) = csv_data_obj.csv_list(csv_data_obj.title_row, now_orderdetail_title_clm)
                Next
            Else
                If IsArray(middleHeaders) Then
                    ' Ensure the array has at least one element before using UBound/LBound
                    On Error Resume Next
                    Dim mhLBound As Long, mhUBound As Long
                    mhLBound = LBound(middleHeaders)
                    mhUBound = UBound(middleHeaders)
                    If Err.Number <> 0 Then
                        Err.Clear
                    Else
                        If mhUBound >= mhLBound Then
                            For now_orderdetail_title_clm = 1 To mhUBound - mhLBound + 1
                                sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, now_orderdetail_title_clm) = middleHeaders(now_orderdetail_title_clm - 1 + mhLBound)
                            Next
                        End If
                    End If
                    On Error GoTo 0
                Else
                    ' fallback to csv header
                    For now_orderdetail_title_clm = 1 To UBound(csv_data_obj.csv_list, 2)
                        sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, now_orderdetail_title_clm) = csv_data_obj.csv_list(csv_data_obj.title_row, now_orderdetail_title_clm)
                    Next
                End If
            End If
            '色に応じて配列に振り分ける
            '注: シート全体の行色ではなく、A～AH(1～34) と CW(101) 列のセル色を個別に確認して判定する
            For now_w2p_row = 2 To UBound(w2p_data.w2p_list, 1)
                Dim matched As Boolean
                matched = False
                Dim anyColored As Boolean
                anyColored = False
                Dim cellColor As Long
                ' まず A～AH (1～34) を順に確認
                For checkClm = 1 To 34
                    cellColor = .Cells(now_w2p_row, checkClm).Interior.Color
                    If .Cells(now_w2p_row, checkClm).Interior.ColorIndex <> syokika_color Then anyColored = True
                    If cellColor = patern_1_color Then
                        sep_w2p_data.shindou_list_count = sep_w2p_data.shindou_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                        Exit For
                    ElseIf cellColor = patern_2_color Then
                        sep_w2p_data.maru_kyoten_list_count = sep_w2p_data.maru_kyoten_list_count + 1
                        '中間ファイル用配列: if headerMap exists, fill columns according to template order; otherwise copy CSV order
                        If (Not Not headerMap) Then
                            For now_orderdetail_clm = 1 To UBound(headerMap)
                                Dim srcIdx As Long
                                srcIdx = headerMap(now_orderdetail_clm)
                                Dim srcVal As String
                                srcVal = ""
                                If srcIdx >= 1 And srcIdx <= UBound(csv_data_obj.csv_list, 2) Then
                                    srcVal = CStr(csv_data_obj.csv_list(now_w2p_row, srcIdx))
                                End If
                                ' If this template column corresponds to 本発注日時 and value looks like approval code, shift right
                                Dim isHonHacchu As Boolean
                                isHonHacchu = False
                                On Error Resume Next
                                If LBound(expectedHeaders) <= (now_orderdetail_clm - 1) And (now_orderdetail_clm - 1) <= UBound(expectedHeaders) Then
                                    If Trim$(CStr(expectedHeaders(now_orderdetail_clm - 1))) = "本発注日時" Then isHonHacchu = True
                                End If
                                On Error GoTo 0
                                If isHonHacchu Then
                                    Dim vTrim As String
                                    vTrim = Replace(Replace(srcVal, Chr(34), ""), " ", "")
                                    If Len(vTrim) > 0 And Left(vTrim, 1) = "B" And IsNumeric(Mid(vTrim, 2)) Then
                                        ' move to next column if available
                                        If now_orderdetail_clm + 1 <= UBound(sep_w2p_data.maru_middle_list, 2) Then
                                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = ""
                                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm + 1) = vTrim
                                        Else
                                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = srcVal
                                        End If
                                    Else
                                        sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = srcVal
                                    End If
                                Else
                                    sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = srcVal
                                End If
                            Next
                        Else
                            For now_orderdetail_clm = 1 To UBound(csv_data_obj.csv_list, 2)
                                sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = csv_data_obj.csv_list(now_w2p_row, now_orderdetail_clm)
                            Next
                        End If
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.maru_kyoten_list(sep_w2p_data.maru_kyoten_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                        Exit For
                    ElseIf cellColor = patern_3_color Then
                        sep_w2p_data.maru_list_count = sep_w2p_data.maru_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.maru_list(sep_w2p_data.maru_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                        Exit For
                    ElseIf cellColor = patern_4_color Then
                        sep_w2p_data.teikan_list_count = sep_w2p_data.teikan_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.teikan_list(sep_w2p_data.teikan_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                        Exit For
                    End If
                Next
                '色判定はじまり
                If Not matched Then
                    ' A～AH で該当色が無ければ CW 列(101) を確認
                    cellColor = .Cells(now_w2p_row, 101).Interior.Color
                    If .Cells(now_w2p_row, 101).Interior.ColorIndex <> syokika_color Then anyColored = True
                    If cellColor = patern_1_color Then
                        sep_w2p_data.shindou_list_count = sep_w2p_data.shindou_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next

                        matched = True
                    ElseIf cellColor = patern_2_color Then
                        sep_w2p_data.maru_kyoten_list_count = sep_w2p_data.maru_kyoten_list_count + 1
                        ' respect headerMap if available; otherwise copy CSV order
                        If (Not Not headerMap) Then
                            For now_orderdetail_clm = 1 To UBound(headerMap)
                                Dim srcIdxCW As Long
                                srcIdxCW = headerMap(now_orderdetail_clm)
                                Dim srcValCW As String
                                srcValCW = ""
                                If srcIdxCW >= 1 And srcIdxCW <= UBound(csv_data_obj.csv_list, 2) Then
                                    srcValCW = CStr(csv_data_obj.csv_list(now_w2p_row, srcIdxCW))
                                End If
                                ' If this template column corresponds to 本発注日時 and value looks like approval code, shift right
                                Dim isHonHacchuCW As Boolean
                                isHonHacchuCW = False
                                On Error Resume Next
                                If LBound(expectedHeaders) <= (now_orderdetail_clm - 1) And (now_orderdetail_clm - 1) <= UBound(expectedHeaders) Then
                                    If Trim$(CStr(expectedHeaders(now_orderdetail_clm - 1))) = "本発注日時" Then isHonHacchuCW = True
                                End If
                                On Error GoTo 0
                                If isHonHacchuCW Then
                                    Dim vTrimCW As String
                                    vTrimCW = Replace(Replace(srcValCW, Chr(34), ""), " ", "")
                                    If Len(vTrimCW) > 0 And Left(vTrimCW, 1) = "B" And IsNumeric(Mid(vTrimCW, 2)) Then
                                        If now_orderdetail_clm + 1 <= UBound(sep_w2p_data.maru_middle_list, 2) Then
                                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = ""
                                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm + 1) = vTrimCW
                                        Else
                                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = srcValCW
                                        End If
                                    Else
                                        sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = srcValCW
                                    End If
                                Else
                                    sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = srcValCW
                                End If
                            Next
                        Else
                            For now_orderdetail_clm = 1 To UBound(csv_data_obj.csv_list, 2)
                                sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_count, now_orderdetail_clm) = csv_data_obj.csv_list(now_w2p_row, now_orderdetail_clm)
                            Next
                        End If
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.maru_kyoten_list(sep_w2p_data.maru_kyoten_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                    ElseIf cellColor = patern_3_color Then
                        sep_w2p_data.maru_list_count = sep_w2p_data.maru_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.maru_list(sep_w2p_data.maru_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                    ElseIf cellColor = patern_4_color Then
                        sep_w2p_data.teikan_list_count = sep_w2p_data.teikan_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.teikan_list(sep_w2p_data.teikan_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                    ElseIf anyColored Then
                        ' 色が設定されているがパターン色ではない場合はパターン1(新藤)扱いにする既存の挙動に合わせる
                        sep_w2p_data.shindou_list_count = sep_w2p_data.shindou_list_count + 1
                        For now_w2p_clm = 1 To UBound(w2p_data.w2p_list, 2)
                            sep_w2p_data.shindou_list(sep_w2p_data.shindou_list_count, now_w2p_clm) = w2p_data.w2p_list(now_w2p_row, now_w2p_clm)
                        Next
                        matched = True
                    End If
                End If
                '色判定おわり
            Next
            
            
            '【受注データ csv】フォルダ作成
            Dim cre_folder As String
            cre_folder = my_dir & "\" & csv_folder_name
            '既に存在している場合作成しない
            If Dir(cre_folder, vbDirectory) = "" Then
                MkDir cre_folder
            End If
            
            '定款ファイル名のYYYYMMDDを日付に変換
            Dim teikan_filename As String
            teikan_filename = set_file_name_data.file_name_list(set_file_name_data.teikan_list_row, set_file_name_data.file_name_clm)
            teikan_filename = Replace(teikan_filename, "YYYYMMDD", Format(Now, "yyyymmdd"))
            teikan_filename = Replace(teikan_filename, "YYMMDD", Format(Now, "yymmdd"))
            
            '出力データが１件以上ある場合のみ書き出し
            If sep_w2p_data.teikan_list_count >= 2 Then
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
            
            '出力データが１件以上ある場合のみCSV書き出し
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
            
            '出力データが1件以上ある場合のみCSV書き出し
            If sep_w2p_data.maru_kyoten_list_count >= 2 Then
                'ブックの新規作成:パターン2
                'ファイル名のYYYYMMDDを当日の日付に変換
                Dim file2_name As String
                file2_name = set_file_name_data.file_name_list(set_file_name_data.kyoten_list_row, set_file_name_data.file_name_clm)
                file2_name = Replace(file2_name, "YYYYMMDD", Format(Now, "yyyymmdd"))
                file2_name = Replace(file2_name, "YYMMDD", Format(Now, "yymmdd"))
                Dim new_book2_path As String
                new_book2_path = cre_folder & "\" & file2_name
                Dim line_kyoten As String
                ' ensure title row for maru_middle_list is populated: prefer template headers, fallback to csv header names
                Dim tclm As Long, haveTitle As Boolean
                haveTitle = False
                For tclm = 1 To UBound(sep_w2p_data.maru_middle_list, 2)
                    If Trim(sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, tclm)) <> "" Then haveTitle = True: Exit For
                Next tclm
                If Not haveTitle Then
                    If headerMapAllocated Then
                        ' build title from template headers (middleHeaders) if available
                        If IsArray(middleHeaders) Then
                            Dim mhL2 As Long, mhU2 As Long
                            On Error Resume Next
                            mhL2 = LBound(middleHeaders)
                            If Err.Number <> 0 Then
                                Err.Clear
                                On Error GoTo 0
                                mhL2 = 1: mhU2 = 0
                            Else
                                mhU2 = UBound(middleHeaders)
                                On Error GoTo 0
                            End If
                            For tclm = 1 To UBound(sep_w2p_data.maru_middle_list, 2)
                                Dim tidx As Long
                                tidx = tclm - 1 + mhL2
                                If mhU2 >= mhL2 And tidx >= mhL2 And tidx <= mhU2 Then
                                    sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, tclm) = Replace(middleHeaders(tidx), Chr(34), "")
                                Else
                                    sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, tclm) = ""
                                End If
                            Next tclm
                        Else
                            ' headerMap allocated but no template headers found: fallback to CSV header row
                            For tclm = 1 To Application.WorksheetFunction.Min(UBound(sep_w2p_data.maru_middle_list, 2), UBound(csv_data_obj.csv_list, 2))
                                sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, tclm) = csv_data_obj.csv_list(csv_data_obj.title_row, tclm)
                            Next tclm
                        End If
                    Else
                        ' fallback to csv header row
                        For tclm = 1 To Application.WorksheetFunction.Min(UBound(sep_w2p_data.maru_middle_list, 2), UBound(csv_data_obj.csv_list, 2))
                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, tclm) = csv_data_obj.csv_list(csv_data_obj.title_row, tclm)
                        Next tclm
                    End If
                End If

                ' CSV出力（共通関数を使用）
                Call ExportArrayToCsv(sep_w2p_data.maru_kyoten_list, new_book2_path, True)
                
                ' maru_middle_listのタイトル行をサニタイズ（SOMPOeケア削除）
                    On Error Resume Next
                    Dim ttlCol As Long, tmpVal As String
                    For ttlCol = 1 To UBound(sep_w2p_data.maru_middle_list, 2)
                        tmpVal = CStr(sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, ttlCol))
                        If Len(Trim$(tmpVal)) > 0 Then
                            ' remove surrounding quotes if present
                            tmpVal = Replace(tmpVal, Chr(34), "")
                            ' normalize full-width space to half-width
                            tmpVal = Replace(tmpVal, "　", " ")
                            ' remove any occurrence of SOMPOケア with optional following space
                            tmpVal = Replace(tmpVal, "SOMPOケア　", "")
                            tmpVal = Replace(tmpVal, "SOMPOケア ", "")
                            tmpVal = Replace(tmpVal, "SOMPOケア", "")
                            tmpVal = Trim$(tmpVal)
                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, ttlCol) = tmpVal
                        End If
                    Next ttlCol

                    ' --- DEBUG DUMP: write mapping and sample rows to a debug file for investigation ---
                    On Error Resume Next
                    Dim dbgStm As Object, dbgPath As String, dbgLine As String
                    Set dbgStm = CreateObject("ADODB.stream")
                    dbgStm.charset = "utf-8"
                    dbgStm.Open
                    dbgPath = my_dir & "\maru_middle_debug.txt"
                    dbgLine = "OutputPath: " & new_book2_path & ".csv" & vbCrLf
                    dbgStm.writetext dbgLine, 1
                    dbgLine = "middleCols=" & CStr(middleCols) & ", headerMapAllocated=" & CStr(headerMapAllocated) & vbCrLf
                    dbgStm.writetext dbgLine, 1
                    ' dump middleHeaders if present
                    If IsArray(middleHeaders) Then
                        dbgStm.writetext "middleHeaders:" & vbCrLf, 1
                        Dim ih As Long
                        For ih = LBound(middleHeaders) To Application.WorksheetFunction.Min(UBound(middleHeaders), LBound(middleHeaders) + 49)
                            dbgStm.writetext CStr(ih - LBound(middleHeaders) + 1) & ":" & middleHeaders(ih) & vbCrLf, 1
                        Next ih
                    Else
                        dbgStm.writetext "middleHeaders: <not array>" & vbCrLf, 1
                    End If
                        ' dump CSV header row
                        dbgStm.writetext "csv header row:" & vbCrLf, 1
                        If UBound(csv_data_obj.csv_list, 2) >= 1 Then
                            Dim ch As Long, chStr As String
                            chStr = ""
                            For ch = 1 To Application.WorksheetFunction.Min(UBound(csv_data_obj.csv_list, 2), 100)
                                chStr = chStr & "[" & CStr(ch) & "]" & CStr(csv_data_obj.csv_list(csv_data_obj.title_row, ch)) & ","
                            Next ch
                            dbgStm.writetext chStr & vbCrLf, 1
                        Else
                            dbgStm.writetext "csv header row: <none>" & vbCrLf, 1
                        End If
                        ' dump W2P header row
                        dbgStm.writetext "w2p header row:" & vbCrLf, 1
                        If UBound(w2p_data.w2p_list, 2) >= 1 Then
                            Dim wh As Long, whStr As String
                            whStr = ""
                            For wh = 1 To Application.WorksheetFunction.Min(UBound(w2p_data.w2p_list, 2), 100)
                                whStr = whStr & "[" & CStr(wh) & "]" & CStr(w2p_data.w2p_list(w2p_data.title_row, wh)) & ","
                            Next wh
                            dbgStm.writetext whStr & vbCrLf, 1
                        End If
                        dbgStm.writetext "csvWasMapped=" & CStr(csvWasMapped) & vbCrLf, 1
                    ' dump headerMap if allocated
                    If headerMapAllocated Then
                        dbgStm.writetext "headerMap:" & vbCrLf, 1
                        Dim im As Long
                        For im = 1 To Application.WorksheetFunction.Min(UBound(headerMap), 100)
                            dbgStm.writetext CStr(im) & ":" & CStr(headerMap(im)) & vbCrLf, 1
                        Next im
                    Else
                        dbgStm.writetext "headerMap: <not allocated>" & vbCrLf, 1
                    End If
                    ' dump sample maru_middle_list rows (first 5 rows, up to 30 cols)
                    dbgStm.writetext "maru_middle_list sample:" & vbCrLf, 1
                    Dim cc As Long
                    For rr = 1 To Application.WorksheetFunction.Min(5, UBound(sep_w2p_data.maru_middle_list, 1))
                        Dim rowStr As String
                        rowStr = ""
                        For cc = 1 To Application.WorksheetFunction.Min(30, UBound(sep_w2p_data.maru_middle_list, 2))
                            rowStr = rowStr & "[" & CStr(cc) & "]" & sep_w2p_data.maru_middle_list(rr, cc) & ","
                        Next cc
                        dbgStm.writetext rowStr & vbCrLf, 1
                    Next rr
                    dbgStm.SaveToFile dbgPath, 2
                    dbgStm.Close
                    Set dbgStm = Nothing
                    On Error GoTo 0
                
                '【受注データ csv】フォルダの直下に【拠点用】フォルダ作成
                Dim cre_kyoten_folder As String
                cre_kyoten_folder = cre_folder & "\" & csv_kyoten_folder_name
                If Dir(cre_kyoten_folder, vbDirectory) = "" Then
                    MkDir cre_kyoten_folder
                End If
                
                'ブックの新規作成:中間ファイル
                'ファイル名のYYYYMMDDを当日の日付に変換
                Dim file2_middle_name As String
                file2_middle_name = set_file_name_data.file_name_list(set_file_name_data.kyoten_list_row, set_file_name_data.file_name_clm)
                file2_middle_name = Replace(file2_middle_name, "YYYYMMDD", Format(Now, "yyyymmdd"))
                file2_middle_name = Replace(file2_middle_name, "YYMMDD", Format(Now, "yymmdd"))
                new_book2_path = cre_kyoten_folder & "\" & file2_middle_name & middle_filename
                
                ' SOMPOケア削除処理
                Dim ttlCol2 As Long, tmpVal2 As String
                For ttlCol2 = 1 To UBound(sep_w2p_data.maru_middle_list, 2)
                    tmpVal2 = CStr(sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, ttlCol2))
                    If Len(Trim$(tmpVal2)) > 0 Then
                        tmpVal2 = Replace(tmpVal2, "　", " ")
                        tmpVal2 = Trim$(tmpVal2)
                        If StrComp(tmpVal2, "SOMPOケア", vbBinaryCompare) = 0 Then
                            sep_w2p_data.maru_middle_list(sep_w2p_data.maru_kyoten_list_title_row, ttlCol2) = ""
                        End If
                    End If
                Next ttlCol2
                
                ' CSV出力（共通関数を使用、UTF-8エンコーディング）
                Call ExportArrayToCsv(sep_w2p_data.maru_middle_list, new_book2_path, True, "UTF-8")
            End If
        End With
        MsgBox "完了しました。"
    End If
    
    '「作業指示書作成リスト」シート保護
    ThisWorkbook.Worksheets(shijisyo_list_sheet).Protect
    
    
End Sub

'==============================================
' CSV出力共通処理
' 配列データをCSVファイルに書き出す
'==============================================
Private Sub ExportArrayToCsv(ByRef dataArray As Variant, ByVal outputPath As String, _
                              Optional ByVal clearPriceColumns As Boolean = False, _
                              Optional ByVal charset As String = "Shift-JIS")
    
    Dim line As String
    Dim now_row As Long, now_clm As Long
    
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
