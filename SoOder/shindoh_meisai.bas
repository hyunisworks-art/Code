Public Type shindoh_meisai
    
    list() As Variant
    no_clm As Long
    order_no_clm As Long
    order_date_clm As Long
    order_name_clm As Long
    haisou_name_clm As Long
    syouhin_code_clm As Long
    syouhin_name_clm As Long
    order_amo_clm As Long
    syoukei_clm As Long
    konpou_clm As Long
    bikou_clm As Long
    end_clm As Long
    
End Type

Public Type shime_hed
    
    list() As Variant
    out_hed_row As Long
    in_match_row As Long
    last_row As Long
    last_clm As Long
    
End Type

Public Type souryou_data
    
    list() As Variant
    'patern1_row As Long
    'patern2_row As Long
    'patern3_row As Long
    haisousaki_clm As Long
    'syouhin_code_clm As Long
    syouhin_code_start_clm As Long
    syouhin_code_end_clm As Long
    other_row As Long
    pattern_start_row As Long
    end_row As Long
    end_clm As Long
    now_end_clm As Long
    souryou_clm As Long
    
End Type

Sub meisai_shindoh()
    
    Dim file_names As Variant
    Dim csv_files As Variant
    Dim min_date As Date
    Dim max_date As Date
    Dim file_date As Date
    Dim shindoh_filename As String
    Dim select_filename As String
    
    '統合データ列番号設定
    Dim shindoh As shindoh_meisai
    shindoh.no_clm = 1
    shindoh.order_no_clm = 2
    shindoh.order_date_clm = 3
    shindoh.order_name_clm = 4
    shindoh.haisou_name_clm = 5
    shindoh.syouhin_code_clm = 6
    shindoh.syouhin_name_clm = 7
    shindoh.order_amo_clm = 8
    shindoh.syoukei_clm = 9
    shindoh.konpou_clm = 10
    shindoh.bikou_clm = 11
    shindoh.end_clm = 11
    
    'CSVファイルの列番号設定
    Dim csv_data As csv_data
    csv_data.title_row = 1
    csv_data.store_clm = 1
    csv_data.order_nom_clm = 2
    csv_data.order_date_clm = 4
    csv_data.order_name_clm = 7
    csv_data.order_status_clm = 19
    csv_data.item_code_clm = 21
    csv_data.item_name_clm = 22
    csv_data.item_count_clm = 23
    csv_data.haisou_order_name_clm = 8
    csv_data.haisou_name_clm = 12
    csv_data.haisou_post_clm = 13
    csv_data.haisou_address1_clm = 14
    csv_data.haisou_address2_clm = 15
    csv_data.haisou_address3_clm = 16
    csv_data.haisou_tel_clm = 17
    csv_data.haisousaki_tantou_clm = 18
    csv_data.syouhin_code_clm = 20
    csv_data.syouhin_name_clm = 21
    csv_data.order_amo_clm = 22
    csv_data.kingaku_clm = 38
    
    '【ヘッダー】新藤様用締めデータの内容取得
    Dim shindoh_hed As shime_hed
    With ThisWorkbook.Worksheets(shime_hed_sheet)
        .Unprotect
        With .Cells(1, 1).SpecialCells(xlLastCell)
            shindoh_hed.last_row = .row
            shindoh_hed.last_clm = .Column
        End With
        shindoh_hed.list = .Range(.Cells(1, 1), .Cells(shindoh_hed.last_row, shindoh_hed.last_clm)).Value
        .Protect
    End With
    
    shindoh_hed.out_hed_row = 1
    shindoh_hed.in_match_row = 2
        
    '集計対象のファイルを複数選択
    Dim my_dir As String
    Dim my_file As String
    my_file = ThisWorkbook.path & "\" & ThisWorkbook.Name
    Dim rng_thisworkbook As Range
    Set rng_thisworkbook = ThisWorkbook.Worksheets(file_name_save_sheet).Cells(2, 1)
    my_file = GetFileName(my_file, rng_thisworkbook)
    my_dir = Left(my_file, InStrRev(my_file, "\"))
    
    ChDir (my_dir & "\" & csv_folder_name)
    file_names = Application.GetOpenFilename(FileFilter:="新藤C用データ,*.csv", MultiSelect:=True)
    If Not IsArray(file_names) Then
        Exit Sub
    Else
        '先に集計対象の行数を調べて統合配列の行数を確保しておく
        Dim sum_row As Long
        sum_row = 0
        For Each csv_files In file_names
            Dim csv_sep As Object
            Dim sp_str As String
            Dim sp_str_line As Variant
            Dim line_sep_tab() As Variant
            '指定したCSVファイルをStreamで読み込む
            Set csv_sep = CreateObject("ADODB.stream")
                csv_sep.Charset = "Shift-JIS"
                csv_sep.Open
                csv_sep.LoadFromFile (csv_files)
            sp_str = csv_sep.ReadText
            sp_str_line = Split(sp_str, vbCrLf)
            
            sum_row = sum_row + UBound(sp_str_line) - 1 'ヘッダーを数に入れないため
            
        Next
        
        ReDim shindoh.list(1 To sum_row, 1 To shindoh_hed.last_clm)
        Dim row_count As Long
        row_count = 1
        For Each csv_files In file_names
            
            '複数選択したファイルからファイル名を抽出
            select_filename = Mid(csv_files, InStrRev(csv_files, "\") + 1, Len(csv_files) - InStrRev(csv_files, "\"))
            
            If Right(select_filename, Len(".csv")) <> ".csv" Then
                MsgBox "CSVファイル以外が選択されています。"
                End
            End If
            
            select_filename = Left(select_filename, Len(select_filename) - Len(".csv")) '".csv"削除
            
            '抽出したファイル名から日付のみを抽出
            Dim str_count As Long
            str_count = 0
            Dim match_str As Long
            match_str = 0
            Dim word() As String
            Dim date_flg As Boolean
            date_flg = False
            Dim date_len As Long
            date_len = 8
            ReDim word(1 To Len(select_filename))
            
            For now_str = 1 To Len(select_filename)
                word(now_str) = Mid(select_filename, now_str, 1)
            Next
            
            For str_now = 1 To UBound(word)
                If IsNumeric(word(str_now)) = True Then
                    str_count = str_count + 1
                Else
                    str_count = 0
                End If
                If str_count = date_len Then
                    If str_now = UBound(word) Then
                        match_str = str_now
                        date_flg = True
                        Exit For
                    Else
                        If IsNumeric(word(str_now + 1)) = False Then
                            match_str = str_now
                            date_flg = True
                            Exit For
                        Else
                            str_count = 0
                        End If
                    End If
                End If
            Next
            
            If date_flg = True Then
                file_date = CDate(Format(Mid(select_filename, match_str - date_len + 1, date_len), "####/##/##"))
            End If
            
            If file_date = "0:00:00" Then
                MsgBox "選択したファイルの中に日付が入っていないファイルがあります。"
                End
            End If
            '最も早い日付と遅い日付を取得
            If min_date = "0:00:00" And max_date = "0:00:00" Then
                min_date = file_date
                max_date = file_date
            Else
                If file_date <= min_date Then
                    min_date = file_date
                End If
                If file_date >= max_date Then
                    max_date = file_date
                End If
            End If
            
            '指定したCSVファイルをStreamで読み込む
            Set csv_sep = CreateObject("ADODB.stream")
                csv_sep.Charset = "Shift-JIS"
                csv_sep.Open
                csv_sep.LoadFromFile (csv_files)
            sp_str = csv_sep.ReadText
            sp_str_line = Split(sp_str, vbCrLf)
            
            ReDim csv_data.csv_list(1 To UBound(sp_str_line), 1 To 1)
            
            '読み込んだCSVファイルを2次元配列に格納
            For now_row = 0 To UBound(sp_str_line)
                If InStr(sp_str_line(now_row), ",") <> 0 Then
                    line_sep_tab = splitCsv(sp_str_line(now_row))
                    If now_row = 0 Then
                        ReDim Preserve csv_data.csv_list(1 To UBound(sp_str_line), 1 To UBound(line_sep_tab))
                    End If
                    For now_clm = 1 To UBound(line_sep_tab)
                        csv_data.csv_list(now_row + 1, now_clm) = line_sep_tab(now_clm)
                    Next
                End If
            Next
            
            '読み込んだCSVファイルの該当項目を抽出して、統合配列に入力
            '対応付けに基づいてデータ入力
            For now_row = 2 To UBound(csv_data.csv_list, 1)
                For now_hedclm = LBound(shindoh_hed.list, 2) To UBound(shindoh_hed.list, 2)
                    If shindoh_hed.list(shindoh_hed.in_match_row, now_hedclm) <> "" Then
                        For now_clm = LBound(csv_data.csv_list, 2) To UBound(csv_data.csv_list, 2)
                            If csv_data.csv_list(1, now_clm) = shindoh_hed.list(shindoh_hed.in_match_row, now_hedclm) Then
                                shindoh.list(row_count, now_hedclm) = csv_data.csv_list(now_row, now_clm)
                                Exit For
                            End If
                        Next
                    End If
                Next
                row_count = row_count + 1
            Next
        Next
        
        '「注文日」でソート
        Call sort_result(shindoh.list, LBound(shindoh.list), UBound(shindoh.list), shindoh.order_date_clm)
        
        '「No」列設定
        Dim nom_count As Long
        nom_count = 1

        '「送料振り分け設定シート」の内容取得
        Dim match_flg As Boolean
        Dim sr_data As souryou_data
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(set_souryou_sheet)
        sr_data.other_row = 2
        sr_data.pattern_start_row = 3
        sr_data.end_row = GetLastRow(ws)
        sr_data.souryou_clm = 2
        sr_data.haisousaki_clm = 3
        sr_data.syouhin_code_start_clm = 4
        sr_data.end_clm = GetLastColumn(ws)
        With ws
            '.Unprotect
            sr_data.list = .Range(.Cells(1, 1), .Cells(sr_data.end_row, sr_data.end_clm)).Value
            '.Protect
        End With
        
        '送料 & No 設定
        For now_row = LBound(shindoh.list, 1) To UBound(shindoh.list, 1)
            match_flg = False
            For now_sr_data_row = sr_data.pattern_start_row To UBound(sr_data.list)
                sr_data.now_end_clm = GetLastColumn(ws, now_sr_data_row)
                '配送先名ありパターン
                If sr_data.list(now_sr_data_row, sr_data.haisousaki_clm) <> "" Then
                    If shindoh.list(now_row, shindoh.haisou_name_clm) = sr_data.list(now_sr_data_row, sr_data.haisousaki_clm) Then
                        shindoh.list(now_row, shindoh.konpou_clm) = sr_data.list(now_sr_data_row, sr_data.souryou_clm)
                        match_flg = True
                        Exit For
                    End If
                '配送先名なし＆商品コードありパターン
                ElseIf sr_data.now_end_clm >= sr_data.syouhin_code_start_clm Then
                    For code_clm = sr_data.syouhin_code_start_clm To sr_data.now_end_clm
                        If sr_data.list(now_sr_data_row, code_clm) <> "" Then
                            If CLng(shindoh.list(now_row, shindoh.syouhin_code_clm)) = CLng(sr_data.list(now_sr_data_row, code_clm)) Then
                                shindoh.list(now_row, shindoh.konpou_clm) = sr_data.list(now_sr_data_row, sr_data.souryou_clm)
                                match_flg = True
                                Exit For
                            End If
                        End If
                    Next
                    If match_flg = True Then
                        Exit For
                    End If
                End If
            Next
            If match_flg = False Then
                '送料のみパターン
                If shindoh.list(now_row, shindoh.konpou_clm) = "" Then
                    shindoh.list(now_row, shindoh.konpou_clm) = sr_data.list(sr_data.other_row, sr_data.souryou_clm)
                End If
            End If
            '「No」設定
            shindoh.list(now_row, shindoh.no_clm) = nom_count
            nom_count = nom_count + 1
        Next
        
        'ファイル出力処理
        Dim out_path As String
        out_path = my_dir & "締め処理\【締データ】"
        Dim out_path1 As String
        out_path1 = my_dir & "締め処理"
        
        If Dir(out_path1, vbDirectory) = "" Then
            MkDir out_path1
            MkDir out_path
        ElseIf Dir(out_path, vbDirectory) = "" Then
            MkDir out_path
        End If
        
        ReDim headder(LBound(shindoh_hed.list, 2) To UBound(shindoh_hed.list, 2)) As Variant
        For now_clm = LBound(shindoh_hed.list, 2) To UBound(shindoh_hed.list, 2)
            headder(now_clm) = shindoh_hed.list(shindoh_hed.out_hed_row, now_clm)
        Next
        
        Dim out_bk As Workbook
        Set out_bk = Workbooks.Add
        Dim goukei_row As Long
        goukei_row = UBound(shindoh.list, 1) + 2
        Dim file_name As String
        file_name = "【" & Format(max_date, "mm") & "月分締めデータ】新藤様" & Format(min_date, "yyyymmdd") & "～" & Format(max_date, "yyyymmdd") & ".xlsx"
        
        With out_bk.Worksheets(1)
            '表示形式設定
            .Range(Columns(shindoh.order_amo_clm), Columns(shindoh.konpou_clm)).NumberFormatLocal = "0_ "
            'データ貼り付け
            .Range(.Cells(1, 1), .Cells(1, UBound(headder))).Value = headder
            .Range(.Cells(2, 1), .Cells(UBound(shindoh.list, 1) + 1, UBound(shindoh.list, 2))).Value = shindoh.list
            .Cells(goukei_row, shindoh.syouhin_name_clm).Value = "合計（" & Format(min_date, "mm/dd") & "-" & Format(max_date, "mm/dd") & "）"
            .Cells(goukei_row, shindoh.order_amo_clm).Formula = "=SUM(H" & 2 & ":H" & UBound(shindoh.list, 1) + 1 & ")"
            .Cells(goukei_row, shindoh.syoukei_clm).Formula = "=SUM(I" & 2 & ":I" & UBound(shindoh.list, 1) + 1 & ")"
            .Cells(goukei_row, shindoh.konpou_clm).Formula = "=SUM(J" & 2 & ":J" & UBound(shindoh.list, 1) + 1 & ")"
            .Cells(goukei_row, shindoh.bikou_clm).NumberFormatLocal = "0_ "
            .Cells(goukei_row, shindoh.bikou_clm).Formula = "=SUM(I" & goukei_row & ":J" & goukei_row & ")"
            
            '背景色 & 太字処理
            .Range(.Cells(1, 1), .Cells(1, UBound(headder))).Interior.Color = RGB(220, 230, 241)
            .Cells(goukei_row, shindoh.syoukei_clm).Interior.Color = RGB(220, 230, 241)
            .Cells(goukei_row, shindoh.konpou_clm).Interior.Color = RGB(220, 230, 241)
            .Cells(goukei_row, shindoh.bikou_clm).Interior.ColorIndex = 6
            .Range(.Cells(1, UBound(headder) - 1), .Cells(1, UBound(headder))).Font.Bold = True
            .Range(.Cells(1, 1), .Cells(goukei_row, UBound(headder))).Borders.LineStyle = xlDot
            .Range(.Cells(1, 1), .Cells(goukei_row, UBound(headder))).Borders.Weight = xlHairline
            .Columns(shindoh.order_date_clm).ColumnWidth = 16
        End With
        
        'out_bk.SaveAs out_path & "\" & file_name
        Call SaveAsRetry(out_bk, out_path & "\" & file_name)
        out_bk.Close
        
        MsgBox "完了しました。"
        
    End If
    
End Sub

Function sort_result(out_list As Variant, min_row As Long, max_row As Long, sort_key As Long)

    Dim set_min_row As Double
    Dim set_max_row As Double
    
    Dim row_half As Variant
    Dim have_data As Variant
    
    row_half = CLng(Format(out_list(Int((min_row + max_row) / 2), sort_key), "yyyymmdd")) '日付を数字に変換
    set_min_row = min_row
    set_max_row = max_row
    
    Do
        Do While CLng(Format(out_list(set_min_row, sort_key), "yyyymmdd")) < row_half '日付を数字に変換して比較
            set_min_row = set_min_row + 1
        Loop
        Do While CLng(Format(out_list(set_max_row, sort_key), "yyyymmdd")) > row_half '日付を数字に変換して比較
                set_max_row = set_max_row - 1
        Loop
        If set_min_row >= set_max_row Then
            Exit Do
        End If
        For now_row = LBound(out_list, 2) To UBound(out_list, 2)
           have_data = out_list(set_min_row, now_row)
           out_list(set_min_row, now_row) = out_list(set_max_row, now_row)
           out_list(set_max_row, now_row) = have_data
        Next
        set_min_row = set_min_row + 1
        set_max_row = set_max_row - 1
    Loop

    If (min_row < set_min_row - 1) Then
        Call sort_result(out_list, min_row, set_min_row - 1, sort_key)
    End If
    If (max_row > set_max_row + 1) Then
        Call sort_result(out_list, set_max_row + 1, max_row, sort_key)
    End If
    
End Function