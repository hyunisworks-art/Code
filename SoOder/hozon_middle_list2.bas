Attribute VB_Name = "middle_list"

Public Sub middle_list_out(ByRef kyoten_list_maru As Variant)
'メインルーチンから、sep_w2p_data.maru_kyoten_list（拠点データ）の情報を持ち込む

    '1:「ファイル名保持」シートのA1セルからCSVファイルパスを取得し、csv_data_obj.csv_listに取り込む
    'CSVパスを取得
    With ThisWorkbook.Worksheets(file_name_save_sheet)
        Dim csv_path As String
        csv_path = .Cells(1, 1).Value
    End With

    'CSVファイルの存在確認
    If csv_path = "" Then
        MsgBox "w2pデータが取り込めていません。" & vbCrLf & "先に「w2pデータ取り込み」ボタンを押下してください。"
        Exit Sub
    End If
    
    If Dir(csv_path, vbNormal) = "" Then
        MsgBox "「w2pデータ取り込み」で指定されたcsvファイルが存在しません。" & vbCrLf & "先に「w2pデータ取り込み」ボタンを押下してください。"
        Exit Sub
    End If

    '====================================================================
    ' CSV取り込みブロック
    ' 指定したCSVファイルを読み込む（UTF-8形式対応）
    '====================================================================
    Dim csv_data_obj As csv_data
    Dim csv_sep As Object
    Set csv_sep = CreateObject("ADODB.Stream")
    csv_sep.charset = "utf-8"
    csv_sep.Open
    csv_sep.LoadFromFile csv_path
    Dim sp_str As String
    sp_str = csv_sep.ReadText ' ファイル全体を文字列として取得
    Dim sp_str_line As Variant
    sp_str_line = Split(sp_str, vbCrLf) '改行で行単位に分割
    csv_sep.Close
    Set csv_sep = Nothing 'ADODB.Streamオブジェクト解放

    'CSV格納用の2次元配列を初期化（行数は確定、列数は後で調整）
    ReDim csv_data_obj.csv_list(1 To UBound(sp_str_line) + 1, 1 To 1)
    Dim now_row As Long
    Dim line_sep_tab() As Variant
    '読み込んだCSVファイルを2次元配列（csv_data_obj.csv_list）に格納
    For now_row = 0 To UBound(sp_str_line)
        If InStr(sp_str_line(now_row), ",") <> 0 Then 'カンマが含まれる行のみ処理
            line_sep_tab = splitCsv(sp_str_line(now_row)) 'CSV行を列ごとに分割
            If now_row = 0 Then
                '1行目（ヘッダー行）で列数を確定し、配列サイズを再設定
                ReDim Preserve csv_data_obj.csv_list(1 To UBound(sp_str_line) + 1, 1 To UBound(line_sep_tab))
            End If
            '各列のデータを2次元配列に格納（0始まり→1始まりに変換）
            For now_clm = 1 To UBound(line_sep_tab)
                csv_data_obj.csv_list(now_row + 1, now_clm) = line_sep_tab(now_clm)
            Next
        End If
    Next now_row
'====================================================================
    '2:CSVファイルを読み込み、W2P形式 or Spinno形式を判定
    '（W2P形式=143列、Spinno形式=20列なので、100列を目安に判断）
    Dim rawCsvColCount As Long
    rawCsvColCount = UBound(csv_data_obj.csv_list, 2)
    If rawCsvColCount < 100 Then
        '3:Spinno形式の場合、W2P形式にマッピングする（サブルーチン）
        Call MapSpinnoToW2P(csv_data_obj.csv_list)
    End If
'====================================================================

'4:W2P形式の列構造になったデータを拠点配送リスト（maru_kyoten_list）と比較し、適合する行だけ残す
    
    '最終出力用の中間データ配列を初期化
    Dim sep_w2p_data As sep_w2p_data
    ReDim sep_w2p_data.maru_middle_list(1 To UBound(csv_data_obj.csv_list, 1), 1 To 143) '列数はW2P形式の143列
    
    '中間ファイルのタイトル行を作成
    Dim i As Long
    For i = 1 To UBound(csv_data_obj.csv_list, 2)
        sep_w2p_data.maru_middle_list(1, i) = csv_data_obj.csv_list(1, i)
    Next
    
    sep_w2p_data.maru_list_count = 1 'タイトル行を1として開始（データ行は2から）
    Dim j As Long
    Dim k As Long
    Dim kyoten_key As String
    Dim csv_key As String
    Dim found As Boolean
    'データ行をループ
    sep_w2p_data.maru_list_count = 1 '1行目はヘッダ
    
    For i = 2 To UBound(kyoten_list_maru, 1)
        kyoten_key = CleanCsvVal(kyoten_list_maru(i, 2)) & "_" & CleanCsvVal(kyoten_list_maru(i, 3))
        
        found = False
    
        For j = 2 To UBound(csv_data_obj.csv_list, 1)
            csv_key = csv_data_obj.csv_list(j, 2) & "_" & csv_data_obj.csv_list(j, 3)
    
            If csv_key = kyoten_key Then
                sep_w2p_data.maru_list_count = sep_w2p_data.maru_list_count + 1
    
                For k = 1 To UBound(csv_data_obj.csv_list, 2)
                    sep_w2p_data.maru_middle_list(sep_w2p_data.maru_list_count, k) = csv_data_obj.csv_list(j, k)
                Next k
                found = True
                Exit For
            End If
        Next j
    Next i

    '====================================================================
    '5:中間ファイルとしてCSV出力する
    'データが存在する場合のみCSV出力
    Dim middle_row_count As Long
    middle_row_count = UBound(sep_w2p_data.maru_middle_list, 1)
    If middle_row_count >= 2 Then
        'ファイル名とパスの設定
        Dim my_dir As String
        Dim my_file As String
        my_file = ThisWorkbook.path & "\" & ThisWorkbook.Name
        Dim rng_thisworkbook As Range
        my_dir = Left(my_file, InStrRev(my_file, "\"))

        '【受注データ csv】フォルダ作成
        Dim cre_folder As String
        cre_folder = my_dir & "\" & csv_folder_name
        If Dir(cre_folder, vbDirectory) = "" Then
            MkDir cre_folder
        End If
        
        '【受注データ csv】フォルダの直下に【拠点用】フォルダ作成
        Dim cre_kyoten_folder As String
        cre_kyoten_folder = cre_folder & "\" & csv_kyoten_folder_name
        If Dir(cre_kyoten_folder, vbDirectory) = "" Then
            MkDir cre_kyoten_folder
        End If
        
        '中間ファイル名の生成
        Dim middle_file_name As String
        middle_file_name = "SOMPO受付" & Format(Now, "yyyymmdd") & " マルテックス様_拠点配送_中間ファイル.csv"
        Dim middle_file_path As String
        middle_file_path = cre_kyoten_folder & "\" & middle_file_name
        ' CSV出力（共通関数を使用、UTF-8エンコーディング）
        Call ExportArrayToCsv(sep_w2p_data.maru_middle_list, middle_file_path, False, "UTF-8")
    End If
End Sub

'====================================================================
' Spinno形式からW2P形式へのマッピング処理
'====================================================================
'3:Spinno形式の場合、W2P形式にマッピングする（サブルーチン）
Private Sub MapSpinnoToW2P(ByRef listArr As Variant)
    
    '受けた配列listArrを、別の配列を作ってコピーする
    Dim rawCsv As Variant
    rawCsv = listArr

    'listArrを初期化しW2P形式のサイズを与える
    Dim cMax As Long
    cMax = 143 'W2P形式は143列
    ReDim listArr(1 To UBound(rawCsv, 1), 1 To cMax) 'W2P形式は143列

    'listArrにW2P形式のヘッダーを付与　※W2P形式の標準ヘッダー（143列）を定義
    Dim expectedHeaders As Variant
    Dim expectedHeadersCsv As String
    'ヘッダー行の期待値文字列（区切り文字は「|」、後でSplit関数で配列化）
    expectedHeadersCsv = "ストア|発注番号|明細番号|注文日|顧客|グループ|発注者ID|発注者|発注者ログインID|発注者コード|件名|配送先名|配送先郵便番号|配送先住所1|配送先住所2|配送先住所3|配送先電話番号|配送先担当者名|" & _
                            "注文状態|本発注日時|商品コード|商品名|注文数量|単価|小計|注文小計|消費税|合計（税込）|配送先名(ｱｲﾃﾑ別）|配送先郵便番号(ｱｲﾃﾑ別）|配送先住所1(ｱｲﾃﾑ別）|配送先住所2(ｱｲﾃﾑ別）|配送先住所3(ｱｲﾃﾑ別）|配送先電話番号(ｱｲﾃﾑ別）|配送先担当者名(ｱｲﾃﾑ別）|入稿データ1|入稿データ2|入稿データ3|" & _
                            "入稿データ4|入稿データ5|入稿データ6|入稿データ7|入稿データ8|入稿データ9|入稿データ10|配送業者番号|出荷日|配送備考|対象ストア|略称|JANコード|説明|備考|納期テキスト|納期（標準）|納期（最短）|納期（締め切り時間）|サムネイル|アイテム詳細画像|納品タイプ|印刷版データ|" & _
                            "自由項目名1|自由項目値1|自由項目名2|自由項目値2|自由項目名3|自由項目値3|自由項目名4|自由項目値4|自由項目名5|自由項目値5|自由項目名6|自由項目値6|自由項目名7|自由項目値7|自由項目名8|自由項目値8|自由項目名9|自由項目値9|自由項目名10|自由項目値10|公開|プロダクトコード|プロダクト名|プロダクト分類|説明（商品プロダクト）|デフォルトプロダクション|" & _
                            "パーツ名1|版データ1|パーツ名2|版データ2|パーツ名3|版データ3|パーツ名4|版データ4|パーツ名5|版データ5|仕上がりサイズ|パーツ名1|裁ち1|色1表|色1裏|用紙1|パーツ名2|裁ち2|色2表|色2裏|用紙2|パーツ名3|裁ち3|色3表|色3裏|用紙3|" & _
                            "パーツ名4|裁ち4|色4表|色4裏|用紙4|パーツ名5|裁ち5|色5表|色5裏|用紙5|折り|特記事項（折）|製本|製本部材|製本部材カラー|特記事項（製本）|穴あけ|特記事項（穴あけ）|断裁|特記事項（断裁)|" & _
                            "その他加工1|特記事項（その他加工1)|その他加工2|特記事項（その他加工2)|その他加工3|特記事項（その他加工3)|その他加工4|特記事項（その他加工4)|その他加工5|特記事項（その他加工5)"
    '期待するCSVヘッダー文字列を「|」区切りで配列に変換
    expectedHeaders = Split(expectedHeadersCsv, "|")

    'ヘッダー行を作成（expectedHeaders をそのまま1行目にセット）
    Dim hc As Long
    For hc = 1 To cMax
        listArr(1, hc) = CStr(expectedHeaders(hc - 1))
    Next hc

    'Spinno形式の各列を対応するW2P形式の列位置にマッピング
    Dim rr As Long
    Dim tmpAddr3 As String
    For rr = 2 To UBound(rawCsv, 1)
        If Not IsEmpty(rawCsv(rr, 2)) Then
            listArr(rr, 1) = "SOMPOケア　" & CStr(rawCsv(rr, 1)) 'ストア列：先頭に"SOMPOケア　"を付与
            listArr(rr, 2) = rawCsv(rr, 2) '発注番号列
            listArr(rr, 3) = rawCsv(rr, 3) '明細番号列
            listArr(rr, 4) = rawCsv(rr, 4) '注文日列
            listArr(rr, 5) = rawCsv(rr, 5) '顧客列
            listArr(rr, 6) = rawCsv(rr, 6) 'グループ列
            listArr(rr, 7) = rawCsv(rr, 7) '発注者ID列
            listArr(rr, 8) = rawCsv(rr, 8) '発注者列
            listArr(rr, 9) = rawCsv(rr, 9) '発注者ログインID列
            listArr(rr, 10) = rawCsv(rr, 10) '発注者コード列
            listArr(rr, 12) = rawCsv(rr, 11) '配送先名列
            listArr(rr, 13) = rawCsv(rr, 12) '配送先郵便番号列
            listArr(rr, 14) = rawCsv(rr, 13) '配送先住所1列
            listArr(rr, 15) = rawCsv(rr, 14) '配送先住所2列
            listArr(rr, 16) = rawCsv(rr, 15) '配送先住所3列
            listArr(rr, 17) = rawCsv(rr, 16) '配送先電話番号列
            listArr(rr, 19) = rawCsv(rr, 17) '注文状態列
            listArr(rr, 21) = rawCsv(rr, 18) '商品コード列
            listArr(rr, 22) = rawCsv(rr, 19) '商品名列
            listArr(rr, 23) = rawCsv(rr, 20) '注文数量列
    
            '配送先住所3列の「？」「?」を「-」に置換
            If Not IsEmpty(listArr(rr, 16)) Then
                tmpAddr3 = CStr(listArr(rr, 16))
                tmpAddr3 = Replace(tmpAddr3, "？", "-")
                tmpAddr3 = Replace(tmpAddr3, "?", "-")
                listArr(rr, 16) = tmpAddr3
            End If
            '配送先住所2・3列の結合処理
            Dim a2 As String
            Dim a3 As String
            a2 = CStr(listArr(rr, 15))
            a3 = CStr(listArr(rr, 16))
            If Trim(a3) <> "" Then
                listArr(rr, 15) = Trim(a2 & " " & a3)
                listArr(rr, 16) = ""
            End If
        End If
    Next rr
End Sub
'====================================================================