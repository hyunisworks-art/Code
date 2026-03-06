'このモジュールの内容は、Sheet1にあるフォームのコードだったが、shijisho_fillin.basに移動した。
''将来的に不要になる可能性が高い。（2026-01 西村記載）
Option Explicit
Public Data_Count As Integer    'データ行数カウント
Public Err_Count As Integer     'エラーデータカウント
Public Err_txt As String        'エラーログ格納用
'---------------------------------------------
'作業指示書Excleファイルを作成する
'※処理は1行ずつ行う。
'更新履歴
'2014-06-11 ver3.00　ワタミ介護案件対応
'---------------------------------------------
Public Sub cmd01_Click()
'On Error GoTo err_do
    
    Dim my_file As String
    my_file = ThisWorkbook.path & "\" & ThisWorkbook.Name
    Dim rng_thisworkbook As Range
    Set rng_thisworkbook = ThisWorkbook.Worksheets(file_name_save_sheet).Cells(2, 1)
    my_file = GetFileName(my_file, rng_thisworkbook)
    Dim my_dir As String
    my_dir = Left(my_file, InStrRev(my_file, "\"))
    
    '---------------------------------------------
    '担当者の登録 ※担当変更時はここを変えてください
    '---------------------------------------------
    Const tantou = "山中"

    Dim i, j As Integer
    Dim myseek As Object

    Dim seek_str As String      '検索対象文字
    'Dim Arr_str(38) As String   '情報収集用配列
    'Dim Arr_str(41) As String   '情報収集用配列
    Dim Arr_str(57) As String   '情報収集用配列
    Dim wkshapeflg As String
    
    Dim rngTarget As Range
    
    Dim save_add As String      'ファイル保存先アドレス
    Dim save_name As String     '保存ファイル名

    '雛形ファイルの保存先
    'ver 3.00変更　ファイル名に合わせて変更すること。
    Const temp_file = "作業指示書　兼　記録表マスタｖ3.00.xlt"

    Application.ScreenUpdating = False

    '初期化
    Data_Count = 0
    Err_Count = 0
    Err_txt = ""
    save_add = ""
    save_name = ""
    seek_str = ""
    i = 2
    j = 0

    Worksheets(1).Select

    '行数をカウントする
    Data_Count = Range("B1").CurrentRegion.Rows.count
    
    'データがない場合、何もせずに終了
    If Data_Count = 1 Then
        MsgBox "作業元データが貼り付けられていません。" & vbCrLf & "作業データ貼付けを行ってから、再度実行してください。"
        Exit Sub
    End If

    Do Until i > Data_Count
    '1行ごとに、必要な情報を配列に格納
    If Cells(i, 4).Value = "" Then Exit Do
        Arr_str(1) = Cells(i, 4).Value      '商品コード
        seek_str = Cells(i, 4).Value        '商品コードを検索文字に代入
        'Arr_str(2) = Cells(i, 5).Value      '製品名 '一枚目製品名もマスタから取得するため、コメントアウト
        Arr_str(12) = Trim(Cells(i, 9).Value) & "　担当：" & Trim(Cells(i, 10).Value) & _
                        "様宛　" & "〒" & Trim(Cells(i, 12).Value) & " " & Trim(Cells(i, 13).Value & Cells(i, 14).Value & Cells(i, 15).Value) & _
                        "　" & Trim(Cells(i, 11).Value) '納品先
        Arr_str(13) = Cells(i, 17).Value        '数量
'        Arr_str(14) = Cells(i, 3).Value         '納期
        Arr_str(14) = Cells(i, 26).Value         '納期 ver3.00変更 Z列から取得に変更
        Arr_str(17) = Format(Date, "yyyy/mm/dd")    '作成日
        Arr_str(18) = Cells(i, 20).Value        '受注No
        Arr_str(25) = Cells(i, 26).Value    '納期
        'Ver2.46追加項目
        'Ver2.47(13-07-19)変更項目　作業シートからマスタへ、売上課所と中分類を移す　作業指示書番号は移したセル分詰める
        'Arr_str(39) = Cells(i, 27).Value   '売上課所
        'Arr_str(40) = Cells(i, 28).Value   '中分類
        'Arr_str(41) = Cells(i, 29).Value   '作業指示書番号
        Arr_str(41) = Cells(i, 27).Value   '作業指示書番号
        'ver3.00追加
        Arr_str(42) = Cells(i, 20).Value '受注№


        Worksheets("製品仕様マスタ").Select

        With Worksheets("製品仕様マスタ")
        'マスタに該当の商品コードがあるか検索
        Set rngTarget = .Columns("B:B") 'ver3.00 マスタの検索列を指定
        .Columns("B:B").Select 'ver3.00 検索範囲にカーソルを移動
'        Set myseek = Worksheets("製品仕様マスタ").Columus("B").Find(seek_str)
        'Set myseek = rngTarget.Find(seek_str)
         
        Set myseek = Selection.Find(What:=seek_str _
                            , LookIn:=xlValues _
                          , LookAt:=xlWhole _
                          , SearchOrder:=xlByColumns _
                          , SearchDirection:=xlNext _
                          , MatchCase:=False _
                          , MatchByte:=False _
                          , SearchFormat:=False) '作成リストの商品マスタに一致する値を抽出

        End With

        If myseek Is Nothing Then
            '該当なし
            Err_txt = Err_txt & Cells(i, 4).Value & " は製品仕様マスタに存在しません。" & vbCrLf
            Err_Count = Err_Count + 1
            Cells(i, 4).Interior.ColorIndex = 36    'マスタにない商品コードのセルを黄色にする。
            
        Else
            '該当あり
            j = Worksheets("製品仕様マスタ").Cells.Find(seek_str).row
    
            '製品仕様マスタから商品コードを元に検索し、必要な情報を配列に格納する
            'ver3.00追加
            Arr_str(2) = Worksheets("製品仕様マスタ").Cells(j, 4).Value     '製品名
            
            Arr_str(3) = Worksheets("製品仕様マスタ").Cells(j, 5).Value     '表紙用紙
            Arr_str(4) = Worksheets("製品仕様マスタ").Cells(j, 6).Value     '本文用紙
            Arr_str(5) = Worksheets("製品仕様マスタ").Cells(j, 7).Value     '表紙刷色
            Arr_str(6) = Worksheets("製品仕様マスタ").Cells(j, 8).Value     '本文刷色
            Arr_str(7) = Worksheets("製品仕様マスタ").Cells(j, 11).Value     '製本仕様
            Arr_str(8) = Worksheets("製品仕様マスタ").Cells(j, 9).Value     '仕上りサイズ
            Arr_str(9) = Worksheets("製品仕様マスタ").Cells(j, 10).Value     '本文ページ数
            Arr_str(10) = Worksheets("製品仕様マスタ").Cells(j, 12).Value    'その他仕様
            Arr_str(11) = ""    'サンプル送り先
            Arr_str(15) = Worksheets("製品仕様マスタ").Cells(j, 15).Value     '入稿形態
            Arr_str(16) = Worksheets("製品仕様マスタ").Cells(j, 13).Value     '差込み有無
            Arr_str(19) = Worksheets("製品仕様マスタ").Cells(j, 3).Value     '商品コード2
            
            'Ver2.0追加項目
            Arr_str(20) = Worksheets("製品仕様マスタ").Cells(j, 16).Value    'サンプル
            Arr_str(21) = Worksheets("製品仕様マスタ").Cells(j, 17).Value    '製品分野
            Arr_str(22) = Worksheets("製品仕様マスタ").Cells(j, 18).Value    '裁断加工
            Arr_str(23) = Worksheets("製品仕様マスタ").Cells(j, 19).Value    'パンチ加工
            Arr_str(24) = Worksheets("製品仕様マスタ").Cells(j, 20).Value    '個人情報
            
            '2.41追加項目　依頼先格納
            Arr_str(26) = Worksheets("製品仕様マスタ").Cells(j, 21).Value    '依頼先

            '2.43追加項目　差替特記事項格納
            Arr_str(27) = Worksheets("製品仕様マスタ").Cells(j, 14).Value    '差替特記事項

            '2.44追加項目　Arr_str(16) が無※の場合、PPマーク入りが挿入されます。
            If Arr_str(16) = "無※" Then
                If Arr_str(27) = "" Then
                    Arr_str(27) = "PPマーク入り"
                ElseIf Arr_str(27) = "PPマーク入り" Then
                    Arr_str(27) = "PPマーク入り"
                Else
                    Arr_str(27) = "PPマーク入り　" & Arr_str(27)
                End If
            End If

            'Ver2.46追加項目
            Arr_str(28) = Worksheets("製品仕様マスタ").Cells(j, 22).Value    '価格表_表紙用紙
            Arr_str(29) = Worksheets("製品仕様マスタ").Cells(j, 23).Value    '価格表_本文用紙
            Arr_str(30) = Worksheets("製品仕様マスタ").Cells(j, 24).Value    '価格表_表紙刷色
            Arr_str(31) = Worksheets("製品仕様マスタ").Cells(j, 25).Value    '価格表_本文刷色
            Arr_str(32) = Worksheets("製品仕様マスタ").Cells(j, 26).Value    '価格表_本文ページ数
            Arr_str(33) = Worksheets("製品仕様マスタ").Cells(j, 27).Value    '価格表_製本仕様
            Arr_str(34) = Worksheets("製品仕様マスタ").Cells(j, 28).Value    '価格表_差込有無
            Arr_str(35) = Worksheets("製品仕様マスタ").Cells(j, 29).Value    '価格表_製品分野
            Arr_str(36) = Worksheets("製品仕様マスタ").Cells(j, 30).Value    '価格表_裁断加工
            Arr_str(37) = Worksheets("製品仕様マスタ").Cells(j, 31).Value    '価格表_パンチ加工
            Arr_str(38) = Worksheets("製品仕様マスタ").Cells(j, 32).Value    '価格表_差替特記事項

        'Ver2.47(13-07-19)追加項目
        Arr_str(39) = Worksheets("製品仕様マスタ").Cells(j, 33).Value    '価格表_売上課所
        Arr_str(40) = Worksheets("製品仕様マスタ").Cells(j, 34).Value    '価格表_中分類

           'ver3.00追加項目
        Arr_str(43) = Worksheets("製品仕様マスタ").Cells(j, 35).Value '商品コード（価格表）
        Arr_str(44) = Worksheets("製品仕様マスタ").Cells(j, 36).Value '製品名（価格表）
        Arr_str(45) = Worksheets("製品仕様マスタ").Cells(j, 37).Value '（入居案内）商品ｺｰﾄﾞ
        Arr_str(46) = Worksheets("製品仕様マスタ").Cells(j, 38).Value '(入居案内)製品名
        Arr_str(47) = Worksheets("製品仕様マスタ").Cells(j, 39).Value '(入居案内)表紙用紙
        Arr_str(48) = Worksheets("製品仕様マスタ").Cells(j, 40).Value '(入居案内)本文用紙
        Arr_str(49) = Worksheets("製品仕様マスタ").Cells(j, 41).Value '"(入居案内)表紙刷色"
        Arr_str(50) = Worksheets("製品仕様マスタ").Cells(j, 42).Value '"(入居案内)本文刷色"
        Arr_str(51) = Worksheets("製品仕様マスタ").Cells(j, 43).Value '"(入居案内)本文ﾍﾟｰｼﾞ数"
        Arr_str(52) = Worksheets("製品仕様マスタ").Cells(j, 44).Value '(入居案内)製本仕様"
        Arr_str(53) = Worksheets("製品仕様マスタ").Cells(j, 45).Value '(入居案内)差込有無※"
        Arr_str(54) = Worksheets("製品仕様マスタ").Cells(j, 46).Value '(入居案内)製品分野"
        Arr_str(55) = Worksheets("製品仕様マスタ").Cells(j, 47).Value '(入居案内)裁断加工"
        Arr_str(56) = Worksheets("製品仕様マスタ").Cells(j, 48).Value '(入居案内)パンチ加工"

            'Excleテンプレートを開く
            If Arr_str(26) = "平和島" Then
                save_add = my_dir & "【平和島】" & Format(Date, "yyyymmdd")
            ElseIf Arr_str(26) = "新藤" Then
                save_add = my_dir & "【新藤】" & Format(Date, "yyyymmdd")
            Else
                save_add = my_dir & "【" & Arr_str(26) & "】" & Format(Date, "yyyymmdd")
            End If
 '2013-10-11 savenameに納期を追加
            save_name = "作業指示書　兼　記録表＜" & Arr_str(1) & "_" & Arr_str(18) & "_" & Arr_str(25) & "＞.xls"
            Workbooks.Open (my_dir & temp_file)
            
            '配列のデータを書出し
' v1.0           With Worksheets("量産")
'                .Range("C5").Value = Arr_str(19) '商品コード
'                .Range("C6").Value = Arr_str(2) '製品名
'                .Range("C7").Value = Arr_str(3) '表紙用紙 C7
'                .Range("C8").Value = Arr_str(4) '本文用紙 C8
'                .Range("F5").Value = Arr_str(5) '表紙刷色 F5
'                .Range("H5").Value = Arr_str(6) '本文刷色 H5
'                .Range("C9").Value = Arr_str(7) '製本仕様 C9
'                .Range("H6").Value = Arr_str(8) '仕上りサイズ H6
'                .Range("H7").Value = Arr_str(9) '本文ページ数 H7
'                .Range("C10").Value = Arr_str(10) 'その他仕様 C10
'                .Range("C11").Value = Arr_str(11) 'サンプル送り先 C11
'                .Range("C12").Value = Arr_str(12) '納品先 C12
'                .Range("H8").Value = Arr_str(13) '数量 H8
'                .Range("H9").Value = DateSerial(Left(Arr_str(14), 4), Mid(Arr_str(14), 5, 2), Right(Arr_str(14), 2)) '納期 H9
'                .Range("H10").Value = Arr_str(15) '入稿形態 H10
'                .Range("H11").Value = Arr_str(16) '差込み有無 H11
'                .Range("J10").Value = Arr_str(17) '作成日 J10
'            End With

'ｖ2.0
             With Worksheets("パンフレット")
                .Range("C8").Value = Arr_str(19) '商品コード
                .Range("C72").Value = Arr_str(19) '商品コード
                .Range("C10").Value = Arr_str(2) '製品名
                .Range("C74").Value = Arr_str(2) '製品名
                .Range("C13").Value = Arr_str(3) '表紙用紙 C7
                .Range("C14").Value = Arr_str(4) '本文用紙 C8
                .Range("F13").Value = Arr_str(5) '表紙刷色 F5
                .Range("F14").Value = Arr_str(6) '本文刷色 H5
                .Range("C12").Value = Arr_str(7) '製本仕様 C9
                .Range("C76").Value = Arr_str(7) '製本仕様 C9
                .Range("H17").Value = Arr_str(8) '仕上りサイズ H6
                .Range("H76").Value = Arr_str(8) '仕上りサイズ H6
                .Range("H12").Value = Arr_str(9) '本文ページ数 H7
'                .Range("C10").Value = Arr_str(10) 'その他仕様 C10→不使用
                .Range("C18").Value = Arr_str(11) 'サンプル送り先 C11
                .Range("C79").Value = Arr_str(11) 'サンプル送り先 C11
                .Range("C19").Value = Arr_str(12) '納品先 C12
                .Range("C80").Value = Arr_str(12) '納品先 C12
                .Range("H10").Value = Arr_str(13) '数量 H8
                .Range("H74").Value = Arr_str(13) '数量 H8
                .Range("H8").Value = DateSerial(Left(Arr_str(25), 4), Mid(Arr_str(25), 5, 2), Right(Arr_str(25), 2)) '納期 H9
               ' .Range("H8").Value = Arr_str(25) '納期 H9
                '.Range("H72").Value = DateSerial(Left(Arr_str(14), 4), Mid(Arr_str(14), 5, 2), Right(Arr_str(14), 2)) '納期 H9
                .Range("H16").Value = Arr_str(15) '入稿形態 H10
                .Range("F12").Value = Arr_str(16) '差込み有無 H11
                .Range("F76").Value = Arr_str(16) '差込み有無 H11
                .Range("J15").Value = Arr_str(17) '作成日 J10
                
                'Ver2.46追加項目
                .Range("H14").Value = Arr_str(39) '計上課所
                .Range("H15").Value = Arr_str(40) '中分類
                .Range("J17").Value = Arr_str(41) '作業指示書番号
                'ver3.00追加項目
                .Range("H7").Value = Arr_str(42) '受注№

'                .Range(").Value = Arr_str(20) 'サンプル　未使用
                .Range("C15").Value = Arr_str(21) '製品分野
                .Range("F15").Value = Arr_str(22) '裁断加工
                .Range("H18").Value = Arr_str(23) 'パンチ加工
                .Range("F77").Value = Arr_str(23) 'パンチ加工
                .Range("H77").Value = Arr_str(24) '個人情報
                .Range("C77").Value = Arr_str(24) '個人情報

                .Range("J19").Value = "無" '個人情報有無
                .Range("J13").Value = tantou '担当者欄
                .Range("J10").Value = "リピート" '新規/リピ
                .Range("C17").Value = "リコージャパン㈱" '新規/リピ
                
                '2.43追加項目　差替特記事項格納
                .Range("C16").Value = Arr_str(27) '差替特記事項入力

             End With
            
                'ver3.00追加項目

                

                If Arr_str(51) <> "" And (IsNull(Arr_str(51)) = False) Then
                

                '価格表本文ページ数が入力されている場合、入居案内に同じ内容をコピーする
                        Worksheets("パンフレット").Rows("1:122").Select
                        Selection.Copy
               
                        '入居案内をアクティブにする
                        ActiveWorkbook.Worksheets("入居案内").Activate
                        
            With Worksheets("入居案内")

                        .Range("A1").Select
                        ActiveSheet.Paste
                        
                        .Range("C8").Select
                        .Range("C8").Value = Arr_str(45) '（入居案内）商品ｺｰﾄﾞ
                        .Range("C10").Value = Arr_str(46) '(入居案内)製品名
                        .Range("C12").Value = Arr_str(52)    '(入居案内)製本仕様
                        .Range("C13").Value = Arr_str(47)    '(入居案内)表紙用紙
                        .Range("C14").Value = Arr_str(48)    '(入居案内)本文用紙
                        .Range("C15").Value = Arr_str(54)    '(入居案内)製品分野"
                        .Range("F12").Value = Arr_str(53)    '(入居案内)差込有無※
                        .Range("H12").Value = Arr_str(51)    '(入居案内)本文ﾍﾟｰｼﾞ数
                       .Range("F13").Value = Arr_str(49)    '(入居案内)表紙刷色
                        .Range("F14").Value = Arr_str(50)    '(入居案内)本文刷色
                        .Range("F15").Value = Arr_str(55)    '(入居案内)裁断加工
                        .Range("H18").Value = Arr_str(56)    '(入居案内)パンチ加工
                        
                        .Range("C72").Value = Arr_str(45) '（入居案内）商品ｺｰﾄﾞ
                        .Range("C74").Value = Arr_str(46) '(入居案内)製品名
                        .Range("C76").Value = Arr_str(52)    '(入居案内)製本仕様
                        .Range("C77").Value = Arr_str(54)    '(入居案内)製品分野
                        .Range("F76").Value = Arr_str(53)    '(入居案内)差込有無※
                        .Range("F77").Value = Arr_str(56)    '(入居案内)パンチ加工
                        
                        .Range("B68").Value = "【入居案内】1/2"
                        .Range("B122").Value = "【入居案内】2/2"
                    End With
                
                End If
                
                
                If Arr_str(32) <> "" And (IsNull(Arr_str(32)) = False) Then
                '価格表本文ページ数が入力されている場合、4ページ目に同じ内容をコピーする
                        ActiveWorkbook.Worksheets("入居案内").Activate
                        Worksheets("入居案内").Rows("1:122").Select
                        Selection.Copy
                         'セレクトをアクティブにする
                        ActiveWorkbook.Worksheets("セレクト").Activate
                
                With Worksheets("セレクト")
            
                        .Range("A1").Select
                        ActiveSheet.Paste
                        .Range("C8").Select
                        .Range("C8").Value = Arr_str(43) '商品コード価格表
                        .Range("C10").Value = Arr_str(44) '製品名（価格表）
'                        .Range("C132").Value = "【価格表】 " & .Range("C132").Value
                        '.Range("C196").Value = "【価格表】 " & .Range("C196").Value
                        .Range("C13").Value = Arr_str(28)    '価格表_表紙用紙
                        .Range("C14").Value = Arr_str(29)    '価格表_本文用紙
                        .Range("F13").Value = Arr_str(30)    '価格表_表紙刷色
                        .Range("F14").Value = Arr_str(31)    '価格表_本文刷色
                        .Range("H12").Value = Arr_str(32)    '価格表_本文ページ数
                        '.Range("C260").Value = Arr_str(38)    '価格表_差替特記事項
                        .Range("C12").Value = Arr_str(33)    '価格表_製本仕様
                        .Range("F12").Value = Arr_str(34)    '価格表_差込有無
                        .Range("C15").Value = Arr_str(35)    '価格表_製品分野
                        .Range("F15").Value = Arr_str(36)    '価格表_裁断加工
                        .Range("H18").Value = Arr_str(37)    '価格表_パンチ加工
                        
                        .Range("C72").Value = Arr_str(43) '（入居案内）商品ｺｰﾄﾞ
                        .Range("C74").Value = Arr_str(44) '(入居案内)製品名
                        .Range("C76").Value = Arr_str(33)    '(入居案内)製本仕様
                        .Range("C77").Value = Arr_str(35)    '(入居案内)製品分野
                        .Range("F76").Value = Arr_str(34)    '(入居案内)差込有無※
                        .Range("F77").Value = Arr_str(37)    '(入居案内)パンチ加工
                        
                        .Range("B68").Value = "【セレクト】1/2"
                        .Range("B122").Value = "【セレクト】2/2"
                        
                    End With
                
                End If
                
                
                
                
                If Arr_str(32) = "" And (IsNull(Arr_str(32)) = False) Then
                    '不正行の削除(123：244)
'                        .Rows("123:244").Select
'                        .Rows("245:366").Select
'                        Selection.Delete
'                        ActiveSheet.Shapes("Line 100").Select
'                        Selection.Delete

                        Application.DisplayAlerts = False
                        Worksheets("セレクト").Delete
                        Application.DisplayAlerts = True
                        
                        wkshapeflg = "3"
                End If
    
                If Arr_str(51) = "" And (IsNull(Arr_str(51)) = False) Then
                    '不正行の削除(123：244)
'                        .Rows("123:244").Select
'                        Selection.Delete
                        Application.DisplayAlerts = False
                        Worksheets("入居案内").Delete
                        Application.DisplayAlerts = True
                        
'                        If wkshapeflg <> "1" Then
'                            ActiveSheet.Shapes("Line 100").Select
'                            Selection.Delete
'                        End If
                        wkshapeflg = "2"
                End If
                
            

            '不正列の削除(R:HF)
            ActiveWorkbook.Worksheets("パンフレット").Activate
            With Worksheets("パンフレット")
                .Columns("R:HF").Select
                .Columns("R:HF").Delete
                .Range("A1").Select
            End With
            
            If Arr_str(32) <> "" And (IsNull(Arr_str(32)) = False) Then
                        ActiveWorkbook.Worksheets("セレクト").Activate
                With Worksheets("セレクト")
                    .Columns("R:HF").Select
                    .Columns("R:HF").Delete
                    .Range("A1").Select
                End With
            
            End If
            
            
            If Arr_str(51) <> "" And (IsNull(Arr_str(51)) = False) Then
                        ActiveWorkbook.Worksheets("入居案内").Activate
                With Worksheets("入居案内")
                    .Columns("R:HF").Select
                    .Columns("R:HF").Delete
                    .Range("A1").Select
            
                End With
            
            End If
            
            

            '印刷設定を行う
            Set_Print
            
            'シートの保護を行う
'            Worksheets("パンフレット").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True
        
            '納品日のみ保護を解除する
            ActiveWorkbook.Worksheets("パンフレット").Activate
        With Worksheets("パンフレット")
            .Range("H8:H9").Select
            Selection.Locked = False
'            .Range("H72").MergeArea.Locked = False
            .Range("H72:H73").Locked = False
        End With
        Worksheets("パンフレット").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

        
        If Arr_str(51) <> "" And (IsNull(Arr_str(51)) = False) Then
   
            ActiveWorkbook.Worksheets("入居案内").Activate
            With Worksheets("入居案内")
                .Range("H8:H9").Locked = False
                .Range("H72:H73").Locked = False
            End With
                     Worksheets("入居案内").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
        
        
        If Arr_str(32) <> "" And (IsNull(Arr_str(32)) = False) Then
                    

                        ActiveWorkbook.Worksheets("セレクト").Activate
            With Worksheets("セレクト")
                .Range("H8:H9").Locked = False
                .Range("H72:H73").Locked = False
            End With
            Worksheets("セレクト").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If

            '当日の作業フォルダを作成する
            Make_Folder (save_add)
   
            '重複ファイルを確認し、保存
            If Exist_file(save_add & "\" & save_name) Then
                If Exist_file(save_add & "\" & Left(save_name, Len(save_name) - 4) & "_1.xls") Then
                    'ActiveWorkbook.SaveAs (save_add & "\" & Left(save_name, Len(save_name) - 4) & "_2.xls")
                    Call SaveAsRetry(ActiveWorkbook, save_add & "\" & Left(save_name, Len(save_name) - 4) & "_2.xls")
                    If Exist_file(save_add & "\" & Left(save_name, Len(save_name) - 4) & "_2.xls") Then
                        'ActiveWorkbook.SaveAs (save_add & "\" & Left(save_name, Len(save_name) - 4) & "_3.xls")
                        Call SaveAsRetry(ActiveWorkbook, save_add & "\" & Left(save_name, Len(save_name) - 4) & "_3.xls")
                    Else
                        'ActiveWorkbook.SaveAs (save_add & "\" & Left(save_name, Len(save_name) - 4) & "_2.xls")
                        Call SaveAsRetry(ActiveWorkbook, save_add & "\" & Left(save_name, Len(save_name) - 4) & "_2.xls")
                    End If
                Else
                    'ActiveWorkbook.SaveAs (save_add & "\" & Left(save_name, Len(save_name) - 4) & "_1.xls")
                    Call SaveAsRetry(ActiveWorkbook, save_add & "\" & Left(save_name, Len(save_name) - 4) & "_1.xls")
                End If
            Else
                'ActiveWorkbook.SaveAs (save_add & "\" & save_name)
                Call SaveAsRetry(ActiveWorkbook, save_add & "\" & save_name)
            End If

        ActiveWorkbook.Close
        Erase Arr_str
        End If
        i = i + 1
    Loop
    Worksheets(1).Select
    
    Application.ScreenUpdating = True

    If Err_Count <> 0 Then
        MsgBox Err_txt & vbCrLf & Data_Count - Err_Count - 1 & "件の作業指示書を書き出しました。"
    Else
        MsgBox Data_Count - Err_Count - 1 & "件の作業指示書を書き出しました。"
    End If

    Exit Sub
err_do:
    
'    If Err.Number = "1004" Then
'        MsgBox "「作業指示書　兼　記録表マスタｖ2.1」ファイルが見つかりません。" & vbCrLf & vbCrLf & ActiveWorkbook.Path & "フォルダに" & vbCrLf & _
'                "「作業指示書　兼　記録表マスタｖ2.1.xlt」がある事を確認してください。", vbExclamation, "マスタファイルが見つかりません"
'        Worksheets(1).Select
'        Exit Sub
'    End If
    MsgBox Err.Number & "エラーが発生しました。渡邊までご連絡ください。（´（エ）｀）ノ" & vbCrLf & vbCrLf & "作業は中断されました。ごめんなさい。"
    Exit Sub

End Sub

'---------------------------------------------
'一覧をクリアする
'---------------------------------------------
Private Sub cmd02_Click()
    If MsgBox("リストを全て削除します。よろしいですか？", vbOKCancel) = vbOK Then
        ThisWorkbook.ActiveSheet.Range("A2", Cells(Rows.count, 1).End(xlDown)).EntireRow.Delete
    End If
End Sub

'---------------------------------------------
'プリンタ設定を行う
'---------------------------------------------
Public Function Set_Print()
    With Worksheets("パンフレット")
        .PageSetup.Zoom = 75
        .PageSetup.Orientation = xlPortrait
        .PageSetup.PaperSize = xlPaperA4
    End With
End Function


'---------------------------------------------
'フォルダがない場合にフォルダを作成する。
'---------------------------------------------
Public Function Make_Folder(strDir As String)
Dim myFileSystem As New Scripting.FileSystemObject
    If myFileSystem.FolderExists(strDir) = False Then myFileSystem.CreateFolder strDir
End Function

'---------------------------------------------
'Exist_file_Del(str As String)
'---------------------------------------------
'ファイルがあるかどうかを判別し、
'       引数 = ファイルアドレス（文字列型）
'       戻値 = Y/N型　※エラー判別　Err = False
Public Function Exist_file(strFile As String)
  Dim VarRet As Variant
  Dim myFileSystem As New Scripting.FileSystemObject
  VarRet = myFileSystem.FileExists(strFile)
  If VarRet = True Then
        'ある
        Exist_file = True
  Else
        'ない
        Exist_file = False
  End If
End Function