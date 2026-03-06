Attribute VB_Name = "PublicVars"
' PublicVars.bas
' 公開変数宣言モジュール
Public Const syokika_color As Long = 0 '白色
Public Const color1_R As Long = 220 'パターン1_R要素
Public Const color1_G As Long = 220 'パターン1_G要素
Public Const color1_B As Long = 255 'パターン1_B要素
Public Const color2_R As Long = 190 'パターン2_R要素
Public Const color2_G As Long = 255 'パターン2_G要素
Public Const color2_B As Long = 190 'パターン2_B要素
Public Const color3_R As Long = 255 'パターン3_R要素
Public Const color3_G As Long = 160 'パターン3_G要素
Public Const color3_B As Long = 160 'パターン3_B要素
Public Const color4_R As Long = 255 'パターン4_R要素
Public Const color4_G As Long = 230 'パターン4_G要素
Public Const color4_B As Long = 153 'パターン4_B要素
Public Const shindo_c_nom As Long = 1
Public Const honsya_nom As Long = 2
Public Const kyoten_nom As Long = 3
Public Const print_clm As Long = 9
Public Const file_name_save_sheet As String = "ファイル名保持"
Public Const order_sheet As String = "注文書"
Public Const shijisyo_list_sheet As String = "作業指示書作成リスト"
Public Const w2pdata_sheet As String = "W2Pデータ貼り付け"
Public Const set_syouhin_code_sheet As String = "特定新藤様商品コード設定シート"
Public Const haisousaki_address_sheet As String = "配送先住所"
Public Const set_file_name_sheet As String = "ファイル名設定シート"
Public Const set_souryou_sheet As String = "送料振り分け設定シート"
Public Const order_link_inout_sheet As String = "orderDetailと入庫、出荷の項目の対応付け設定"
Public Const hed_syukka_sheet As String = "【ヘッダー】出荷データ"
Public Const hed_nyuuko_sheet As String = "【ヘッダー】入庫取込データ"
Public Const shime_hed_sheet As String = "【ヘッダー】新藤様用締めデータ"
Public Const teikan_code_sheet As String = "定款コード"
Public Const csv_folder_name As String = "【受注データcsv】"
Public Const teikan_folder As String = "【新藤】預かり"
Public Const csv_kyoten_folder_name As String = "【拠点用】"
Public Const temp_file As String = "作業指示書　兼　記録表マスタｖ3.00.xlt"
Public Const middle_filename As String = "_中間ファイル"

Public Type csv_data
    title_row As Long
    csv_list() As Variant
    store_clm As Long
    order_nom_clm As Long
    order_status_clm As Long
    order_date_clm As Long
    order_name_clm As Long
    item_code_clm As Long
    item_name_clm As Long
    item_count_clm As Long
    haisou_order_name_clm As Long
    haisou_name_clm As Long
    haisou_post_clm As Long
    haisou_address1_clm As Long
    haisou_address2_clm As Long
    haisou_address3_clm As Long
    haisou_tel_clm As Long
    haisousaki_tantou_clm As Long
    haisou_name_item_clm As Long
    haisou_tantou_item_clm As Long
    syouhin_code_clm As Long
    syouhin_name_clm As Long
    order_amo_clm As Long
    kingaku_clm As Long
    
End Type

Public Type w2p_data
    title_row As Long
    w2p_list() As Variant
    store_clm As Long
    order_nom_clm As Long
    order_date_clm As Long
    order_status_clm As Long
    item_code_clm As Long
    item_name_clm As Long
    item_count_clm As Long
    haisou_order_name_clm As Long
    haisou_name_clm As Long
    haisou_post_clm As Long
    haisou_address1_clm As Long
    haisou_address2_clm As Long
    haisou_address3_clm As Long
    haisou_tel_clm As Long
    haisousaki_tantou_clm As Long
    haisou_name_item_clm As Long
    haisou_tantou_item_clm As Long
    nouki_clm As Long
    sagyou_shiji_clm As Long
    syukko_yotei_clm As Long
    end_data_clm As Long
    end_data_row As Long
    
End Type

Public Type haisousaki_data
    address_data() As Variant
    patern_key_row As Long
    title_row As Long
    nomber_clm As Long
    
End Type

Public Type clm_link_data
    title_row As Long
    clm_link_list() As Variant
    order_detail_row As Long
    haisou_address_row As Long
    w2p_data_row As Long
    nyuuko_data_row As Long
    syukka_data_row As Long
    
End Type

Public Type shijisyo_list_data
    end_row As Long '終了行
    end_clm As Long '終了列
    shijisyo_list() As Variant '作業指示書作成リスト用配列
    shijisyo_title_row As Long 'タイトル行
    item_code_clm As Long '商品コード列
    item_name_clm As Long '商品名列
    order_name_clm As Long '注文者名列
    haisou_name_clm As Long '配送先名列
    haisou_tanou_clm As Long '配送先担当者列
    haisou_tel_clm As Long '配送先電話番号列
    haisou_post_clm As Long '配送先郵便番号列
    haisou_address1_clm As Long '配送先住所1列
    haisou_address2_clm As Long '配送先住所2列
    haisou_address3_clm As Long '配送先住所3列
    item_count_clm As Long '数量列
    order_nom_clm As Long '受注番号列
    nouki_clm As Long '納期列
    sagyou_shiji_clm As Long '作業指示列
    syukko_yotei_clm As Long '出庫予定列

End Type

Public Type sep_w2p_data
    shindou_list() As Variant
    maru_kyoten_list() As Variant
    maru_middle_list() As Variant
    maru_list() As Variant
    teikan_list() As Variant
    shindou_list_count As Long
    maru_kyoten_list_count As Long
    maru_list_count As Long
    teikan_list_count As Long
    shindou_list_title_row As Long
    maru_kyoten_list_title_row As Long
    maru_list_title_row As Long
    teikan_list_title_row As Long
    
End Type

Public Type set_file_name_data
    start_row As Long
    start_clm As Long
    file_name_list() As Variant
    file_clm As Long
    file_name_clm As Long
    order_list_row As Long
    shindou_list_row As Long
    kyoten_list_row As Long
    maru_list_row As Long
    nyuuko_list_row As Long
    syukka_list_row As Long
    teikan_list_row As Long
    end_row As Long
    end_clm As Long
    
End Type

Public Type teikan_code_data
    s_row As Long
    code_clm As Long
    end_row As Long
    end_clm As Long
    list() As Variant
    out_list() As Variant
    out_count As Long
    out_title_row As Long
    flg As Boolean
    
End Type

Public Type color_send_days_data
    list() As Variant
    start_row As Long
    end_row As Long
    clm_min_code As Long
    clm_max_code As Long
    clm_color As Long
    clm_to_send_days As Long
    now_min_code As Long
    now_max_code As Long
    
End Type

'値の整形： " を除去、末尾の , を1つだけ除去
Public Function CleanCsvVal(ByVal v As Variant) As String
    Dim s As String
    s = Replace$(CStr(v), """", "")
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)
    CleanCsvVal = s
End Function