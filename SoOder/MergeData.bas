Public Const name_pack_category = "梱包数"
Public Const name_pack_item = "梱包数(個別)"


Sub createOthersideData()
    
    Dim my_dir As String
    Dim my_file As String
    Dim row_last As Long
    Dim clm_last As Long
    Dim data_titles As Variant
    Dim data_order As Variant
    Dim clm_item As Long
    Dim clm_post As Long
    Dim clm_st_address As Long
    Dim clm_end_address As Long
    Dim idx_clm As Long
    Dim clm_day As Long
    Dim clm_count As Long
    Dim csv_file As Variant
    Dim csv_data As Variant
    Dim flg_head As Boolean
    Dim csv_line As Variant
    Dim csv_text As Variant
    Dim csv_item As String
    Dim csv_address As String
    Dim num_address As Variant
    Dim flg_match As Boolean
    Dim idx_tmp As Long
    Dim pre_address As String
    Dim csv_day As Variant
    Dim pre_day As Variant
    Dim add_count As Long
    Dim idx_text As Long
    Dim data_category As Variant
    Dim data_pack As Variant
    Dim idx_row As Long
    Dim tmp_item As String
    Dim flg_packed As Boolean
    Dim idx_pack As Long
    Dim idx_cat As Long
    Dim head_in As Variant
    Dim head_out As Variant
    Dim text_in As String
    Dim cat_head As String
    Dim add_head As String
    Dim idx_clm_title As Long
    Dim text_title As String
    Dim idx_detail_row As Long
    Dim in_row As Long
    Dim now_row As Long
    Dim text_out As String
    Dim out_row As Long
    Dim wbin As Workbook
    Dim wbout As Workbook
    Dim csv_count_val As Long
    Dim tmp_count_val As Long
    my_file = ThisWorkbook.path & "\" & ThisWorkbook.Name
    Dim rng_thisworkbook As Range
    Set rng_thisworkbook = ThisWorkbook.Worksheets(file_name_save_sheet).Cells(2, 1)
    my_file = GetFileName(my_file, rng_thisworkbook)
    my_dir = Left(my_file, InStrRev(my_file, "\"))
    
    Dim file_names As Variant
    If Dir(my_dir & csv_folder_name & "\" & csv_kyoten_folder_name, vbDirectory) = "" Then
        MsgBox "「拠点用」フォルダが見つかりません。" & vbCrLf & "先に「w2pデータ取り込み」と「作業指示書作成」ボタンを押下してください。"
        End
    Else
        ChDir (my_dir & csv_folder_name & "\" & csv_kyoten_folder_name)
        file_names = Application.GetOpenFilename(FileFilter:="拠点データ,*.csv", MultiSelect:=True)
        If Not IsArray(file_names) Then
            Exit Sub
        Else
            With ThisWorkbook
                '対応付けを取得
                With .Worksheets(order_link_inout_sheet)
                    .Unprotect
                    With .Cells(1, 1).SpecialCells(xlLastCell)
                        row_last = .row
                        clm_last = .Column
                    End With
                    data_titles = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                    .Protect
                    
                    'orderDetailシートの列形式を取得
                    data_order = .Range(.Cells(1, 1), .Cells(1, clm_last)).Value
                End With
                
                '商品コード列の指定（CSVの列20）
                clm_item = 20
                
                '郵便番号列の指定(orderDetailの内容は変わらないため、固定値)
                clm_post = 13
                
                '配送先拠点列の指定(orderDetailの内容は変わらないため、固定値)
                Dim clm_address()
                clm_st_address = 12
                clm_end_address = 16
                ReDim clm_address(1 To 1)
                For idx_clm = clm_st_address To clm_end_address
                    '郵便番号以外の列を指定
                    If idx_clm <> clm_post Then
                        clm_address(UBound(clm_address)) = idx_clm
                        ReDim Preserve clm_address(1 To UBound(clm_address) + 1)
                    End If
                Next idx_clm
                ReDim Preserve clm_address(1 To UBound(clm_address) - 1)
                
                '発注日列の指定(orderDetailの内容は変わらないため、固定値)
                clm_day = 4
                
                '注文数量列の指定（CSVの列22）
                clm_count = 22
                
                Dim tmp_detail()
                ReDim tmp_detail(1 To UBound(data_order, 2), 1 To 1)
                
                '各CSVファイルから内容を取得
                For Each csv_file In file_names
                    csv_data = getTextData(csv_file, "utf-8", vbCrLf)
                    
                        
                    flg_head = True
                    
                    '各行から、テキスト情報を取得
                    For Each csv_line In csv_data
                        If flg_head = True Then
                            '先頭行は無視
                            flg_head = False
                        Else
                            csv_text = splitCsv(csv_line)
                            
                            '商品番号と住所の組み合わせを取得
                            If clm_item <= UBound(csv_text) Then
                                csv_item = csv_text(clm_item)
                            Else
                                csv_item = ""
                            End If
                            csv_address = ""
                            For Each num_address In clm_address
                                If num_address <= UBound(csv_text) Then
                                    csv_address = csv_address & "$$$" & csv_text(num_address)
                                End If
                            Next num_address
                            
                            flg_match = False
                            '作成済みのデータと比較して、一致するものがあるか確認
                            For idx_tmp = LBound(tmp_detail, 2) To UBound(tmp_detail, 2)
                                If csv_item = tmp_detail(clm_item, idx_tmp) Then
                                    pre_address = ""
                                    For Each num_address In clm_address
                                        pre_address = pre_address & "$$$" & tmp_detail(num_address, idx_tmp)
                                    Next num_address
                                    If csv_address = pre_address Then
                                        flg_match = True
                                        '作成済みデータと一致するものがあった場合、日付を比較して、早い(古い)方へ数を合算する
                                        If clm_day <= UBound(csv_text) Then
                                            csv_day = csv_text(clm_day)
                                        Else
                                            csv_day = Now
                                        End If
                                        pre_day = tmp_detail(clm_day, idx_tmp)
                                        
                                        '数量を合算（配列の範囲チェックを追加）
                                        If clm_count >= LBound(csv_text) And clm_count <= UBound(csv_text) Then
                                            If IsNumeric(csv_text(clm_count)) And csv_text(clm_count) <> "" Then
                                                csv_count_val = CLng(csv_text(clm_count))
                                            Else
                                                csv_count_val = 0
                                            End If
                                        Else
                                            csv_count_val = 0
                                        End If
                                        
                                        If clm_count >= LBound(tmp_detail, 1) And clm_count <= UBound(tmp_detail, 1) Then
                                            If IsNumeric(tmp_detail(clm_count, idx_tmp)) And tmp_detail(clm_count, idx_tmp) <> "" Then
                                                tmp_count_val = CLng(tmp_detail(clm_count, idx_tmp))
                                            Else
                                                tmp_count_val = 0
                                            End If
                                        Else
                                            tmp_count_val = 0
                                        End If
                                        
                                        add_count = csv_count_val + tmp_count_val
                                        
                                        If CDate(csv_day) < CDate(pre_day) Then
                                            '作成済みデータの方が新しいならば、すべてのデータを新規行で更新
                                            For idx_text = LBound(tmp_detail, 1) To UBound(tmp_detail, 1)
                                                If idx_text >= LBound(csv_text) And idx_text <= UBound(csv_text) Then
                                                    tmp_detail(idx_text, idx_tmp) = csv_text(idx_text)
                                                Else
                                                    tmp_detail(idx_text, idx_tmp) = ""
                                                End If
                                            Next idx_text
                                        End If
                                        '合算値を更新
                                        tmp_detail(clm_count, idx_tmp) = add_count
                                        
                                        Exit For
                                    End If
                                End If
                            Next idx_tmp
                            
                            'もし作成済みデータと一致するものがないならば、新規行を追加
                            If flg_match = False Then
                                For idx_text = LBound(tmp_detail, 1) To UBound(tmp_detail, 1)
                                    If idx_text >= LBound(csv_text) And idx_text <= UBound(csv_text) Then
                                        tmp_detail(idx_text, UBound(tmp_detail, 2)) = csv_text(idx_text)
                                    Else
                                        tmp_detail(idx_text, UBound(tmp_detail, 2)) = ""
                                    End If
                                Next idx_text
                                ReDim Preserve tmp_detail(1 To UBound(data_order, 2), 1 To UBound(tmp_detail, 2) + 1)
                            End If
                        End If
                    Next csv_line
                Next csv_file
                If UBound(tmp_detail, 2) <= 1 Then
                    MsgBox "指定されたCSVの内容が取得できませんでした。"
                    Exit Sub
                Else
                    ReDim Preserve tmp_detail(1 To UBound(data_order, 2), 1 To UBound(tmp_detail, 2) - 1)
                    
                    '梱包の情報を取得
                    With .Worksheets(name_pack_category)
                        .Unprotect
                        With .Cells(1, 1).SpecialCells(xlLastCell)
                            row_last = .row
                            clm_last = .Column
                        End With
                        data_category = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                        .Protect
                    End With
                    
                    '個別梱包の情報を取得
                    With .Worksheets(name_pack_item)
                        .Unprotect
                        With .Cells(1, 1).SpecialCells(xlLastCell)
                            row_last = .row
                            clm_last = .Column
                        End With
                        data_pack = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                        .Protect
                    End With
                    
                    '各内容を確認し、梱包数に応じた除算を行う
                    For idx_row = LBound(tmp_detail, 2) To UBound(tmp_detail, 2)
                        '商品コード取得
                        tmp_item = tmp_detail(clm_item, idx_row)
                        
                        flg_packed = False
                        For idx_pack = LBound(data_pack, 1) + 1 To UBound(data_pack, 1)
                            If tmp_item = data_pack(idx_pack, 1) Then
                                '個別商品コードが一致する場合、対応する梱包数で梱包
                                flg_packed = True
                                
                                If IsNumeric(data_pack(idx_pack, 2)) And data_pack(idx_pack, 2) > 0 Then
                                    If IsNumeric(tmp_detail(clm_count, idx_row)) Then
                                        If tmp_detail(clm_count, idx_row) / data_pack(idx_pack, 2) >= 1 Then
                                            tmp_detail(clm_count, idx_row) = tmp_detail(clm_count, idx_row) / data_pack(idx_pack, 2)
                                        End If
                                    End If
                                End If
                                Exit For
                            End If
                        Next idx_pack
                        
                        '個別梱包が行われなかった場合
                        If flg_packed = False Then
                            For idx_cat = LBound(data_category, 1) + 1 To UBound(data_category, 1)
                                If Left(tmp_item, 1) = Left(data_category(idx_cat, 1), 1) Then
                                    '商品コードが一致する場合、対応する梱包数で梱包
                                    flg_packed = True
                                    If IsNumeric(data_category(idx_cat, 2)) And data_category(idx_cat, 2) > 0 Then
                                        If IsNumeric(tmp_detail(clm_count, idx_row)) Then
                                            If tmp_detail(clm_count, idx_row) / data_category(idx_cat, 2) >= 1 Then
                                                tmp_detail(clm_count, idx_row) = tmp_detail(clm_count, idx_row) / data_category(idx_cat, 2)
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next idx_cat
                        End If
                        
                    Next idx_row
                End If
                
                'ここまでで、orderDetailの内容を統合/梱包
                
                
                
                '入庫、出荷の列形式を取得
                Dim data_in()
                With .Worksheets(hed_nyuuko_sheet)
                    .Unprotect
                    With .Cells(1, 1).SpecialCells(xlLastCell)
                        row_last = .row
                        clm_last = .Column
                    End With
                    head_in = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                    row_last = UBound(tmp_detail, 2) + 1
                    data_in = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                    .Protect
                End With
                
                Dim data_out()
                With .Worksheets(hed_syukka_sheet)
                    .Unprotect
                    With .Cells(1, 1).SpecialCells(xlLastCell)
                        row_last = .row
                        clm_last = .Column
                    End With
                    head_out = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                    row_last = UBound(tmp_detail, 2) + 1
                    data_out = .Range(.Cells(1, 1), .Cells(row_last, clm_last)).Value
                    .Protect
                End With
                
                '入庫取込データの先頭行と、対応するorderDetailの先頭行を取得
                Dim match_flg As Boolean
                match_flg = False
                row_in = 2
                For idx_clm = 1 To UBound(data_in, 2)
                    match_flg = False
                    text_in = data_in(1, idx_clm)
                    cat_head = head_in(2, idx_clm)
                    add_head = head_in(3, idx_clm)
                    
                    For idx_clm_title = 1 To UBound(data_titles, 2)
                        text_title = data_titles(row_in, idx_clm_title)
                        If text_in = text_title Then
                            match_flg = True
                            '対象の列名と、対応付けの列名が一致した場合、orderDetailが基準であるため、その列番号を利用できる
                            For idx_detail_row = LBound(tmp_detail, 2) To UBound(tmp_detail, 2)
                                If cat_head = "" Then
                                    data_in(idx_detail_row + 1, idx_clm) = tmp_detail(idx_clm_title, idx_detail_row)
                                ElseIf cat_head = "固定" Then
                                    data_in(idx_detail_row + 1, idx_clm) = add_head
                                ElseIf cat_head = "日付" Then
                                    data_in(idx_detail_row + 1, idx_clm) = Format(tmp_detail(idx_clm_title, idx_detail_row), add_head)
                                Else
                                    data_in(idx_detail_row + 1, idx_clm) = tmp_detail(idx_clm_title, idx_detail_row)
                                End If
                            Next idx_detail_row
                            Exit For
                        End If
                    Next idx_clm_title
                    '一致する項目がなかった場合
                    If match_flg = False Then
                        For in_row = 2 To UBound(data_in, 1)
                            If cat_head = "固定" Then
                                data_in(in_row, idx_clm) = add_head
                            Else
                                data_in(in_row, idx_clm) = ""
                            End If
                        Next
                    End If
                Next idx_clm
                
                '配送先住所3を配送先住所2に合算
                Dim address2_clm As Long
                address2_clm = 15
                Dim address3_clm As Long
                address3_clm = 16
                For now_row = 1 To UBound(tmp_detail, 2)
                    tmp_detail(address2_clm, now_row) = tmp_detail(address2_clm, now_row) & " " & tmp_detail(address3_clm, now_row)
                    tmp_detail(address3_clm, now_row) = ""
                Next
                
                row_out = 3
                For idx_clm = 1 To UBound(data_out, 2)
                    match_flg = False
                    text_out = data_out(1, idx_clm)
                    cat_head = head_out(2, idx_clm)
                    add_head = head_out(3, idx_clm)
                    
                    For idx_clm_title = 1 To UBound(data_titles, 2)
                        text_title = data_titles(row_out, idx_clm_title)
                        If text_out = text_title Then
                            match_flg = True
                            '対象の列名と、対応付けの列名が一致した場合、orderDetailが基準であるため、その列番号を利用できる
                            For idx_detail_row = LBound(tmp_detail, 2) To UBound(tmp_detail, 2)
                                If cat_head = "" Then
                                    data_out(idx_detail_row + 1, idx_clm) = tmp_detail(idx_clm_title, idx_detail_row)
                                ElseIf cat_head = "固定" Then
                                    data_out(idx_detail_row + 1, idx_clm) = add_head
                                ElseIf cat_head = "日付" Then
                                    data_out(idx_detail_row + 1, idx_clm) = Format(tmp_detail(idx_clm_title, idx_detail_row), add_head)
                                Else
                                    data_out(idx_detail_row + 1, idx_clm) = tmp_detail(idx_clm_title, idx_detail_row)
                                End If
                            Next idx_detail_row
                            Exit For
                        End If
                    Next idx_clm_title
                    '一致する項目がなかった場合
                    If match_flg = False Then
                        For out_row = 2 To UBound(data_out, 1)
                            If cat_head = "固定" Then
                                data_out(out_row, idx_clm) = add_head
                            Else
                                data_out(out_row, idx_clm) = ""
                            End If
                        Next
                    End If
                Next idx_clm
            End With
            
            Set wbin = Workbooks.Add
            Set wbout = Workbooks.Add
            
            With wbin.Worksheets(1)
                .Name = "入庫"
                With .Range(.Cells(1, 1), .Cells(UBound(data_in, 1), UBound(data_in, 2)))
                    .NumberFormatLocal = "@"
                    wbin.Worksheets(1).Columns("N:N").NumberFormatLocal = "0_ "
                    .Value = data_in
                End With
            End With
            With wbout.Worksheets(1)
                .Name = "出荷"
                With .Range(.Cells(1, 1), .Cells(UBound(data_out, 1), UBound(data_out, 2)))
                    .NumberFormatLocal = "@"
                    wbout.Worksheets(1).Columns("AA:AA").NumberFormatLocal = "0_ "
                    .Value = data_out
                End With
            End With
            
            Dim set_file_name_data As set_file_name_data
            With ThisWorkbook.Sheets(set_file_name_sheet)
                set_file_name_data.start_row = 3
                set_file_name_data.start_clm = 1
                set_file_name_data.end_row = 8
                set_file_name_data.end_clm = 2
                set_file_name_data.file_name_clm = 2
                set_file_name_data.nyuuko_list_row = 5
                set_file_name_data.syukka_list_row = 6
                set_file_name_data.file_name_list = .Range(.Cells(set_file_name_data.start_row, set_file_name_data.start_clm), .Cells(set_file_name_data.end_row, set_file_name_data.end_clm)).Value
            End With
            Dim nyuuko_file_name As String
            nyuuko_file_name = Replace(set_file_name_data.file_name_list(set_file_name_data.nyuuko_list_row, set_file_name_data.file_name_clm), "YYYYMMDD", Format(Now, "yyyymmdd"))
            nyuuko_file_name = Replace(nyuuko_file_name, "YYMMDD", Format(Now, "yymmdd"))
            Dim syukka_file_name As String
            syukka_file_name = Replace(set_file_name_data.file_name_list(set_file_name_data.syukka_list_row, set_file_name_data.file_name_clm), "YYYYMMDD", Format(Now, "yyyymmdd"))
            syukka_file_name = Replace(syukka_file_name, "YYMMDD", Format(Now, "yymmdd"))
            
            'Call wbin.SaveAs(fileName:=my_dir & csv_folder_name & "\" & nyuuko_file_name)
            Call SaveAsRetry(wbin, my_dir & csv_folder_name & "\" & nyuuko_file_name & ".xlsx")
            'Call wbout.SaveAs(fileName:=my_dir & csv_folder_name & "\" & syukka_file_name)
            Call SaveAsRetry(wbout, my_dir & csv_folder_name & "\" & syukka_file_name & ".xlsx")
            Application.DisplayAlerts = False
            wbin.Close
            wbout.Close
            Application.DisplayAlerts = True
        End If
    End If
    
    ThisWorkbook.Worksheets(w2pdata_sheet).Activate
    MsgBox "入庫、出荷データを出力しました。"
    
End Sub

Function getTextData(fPath, inputcode, lineSep)
    'fPath:内容を取得するファイル
    'inputcode:取得時のエンコード形式(UTF-8, shift-jis, unicode, euc-jp)
    '**************************
    '指定したファイルの内容を配列に格納する
    '**************************
    Dim inputStream As Object
    Set inputStream = CreateObject("ADODB.Stream")
    inputStream.Open
    inputStream.Type = 2 'adTypeText = 2
    inputStream.Charset = inputcode
    If lineSep = vbLf Then
        inputStream.lineseparator = 10
    End If
    inputStream.LoadFromFile (fPath)
    
    Dim lineText As Long
    lineText = 0
    
    Dim datLine() As String
    Dim txtLine As String
    Dim retText() As String
    
    Do While Not inputStream.EOS
        txtLine = inputStream.ReadText(-2) 'adReadLine = -2
        
        ReDim Preserve retText(lineText)
        retText(lineText) = txtLine
        lineText = lineText + 1
    Loop
    
    inputStream.Close
    Set inputStream = Nothing
    getTextData = retText
End Function