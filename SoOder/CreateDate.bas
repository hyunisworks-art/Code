Public Const name_sheet_area = "配送先都道府県"
Public Const clm_area_name = 1
Public Const clm_area_day = 2

Public Const name_sheet_holyday = "祝日表"
Public Const clm_holyday = 2

Public Const key_numeric = 0
Public Const key_A = 1
Public Const key_B = 2
Public Const key_C = 3

Public Const send_sindoh = 1
Public Const send_sompo = 2
Public Const send_other = 3

Public Const ng_sunday = 1
Public Const ng_sunday_holiday = 2
Public Const ng_all = 3

Public Type RET_DAY
    send_day As Date
    get_day As Date
End Type

Function createDay(rec_day, send_time, key, send_pattern, Optional to_send_day = 3) As RET_DAY
'商品コードの先頭キーと、配送先パターンから、発送日・納品日を算出する
'先に住所から配送期間を取得しておく
    
    send_day = rec_day
    get_day = rec_day
    
    If key = key_numeric Then
        '商品コードに先頭文字がない(数値のみ)パターン
        '発送日が注文日+3営業日(ないし指定があればその営業日数)
        '納品日が上記+配送期間

        '発送日として、3営業日後(ないし指定があればその営業日数後)を取得
        For idx_day = 1 To to_send_day
            send_day = add_holyday(send_day + 1, 1, ng_all)
        Next idx_day
        
        '配送にかかる期間から、納期を算出
        get_day = send_day + send_time
        
        '発送日が日曜だった場合、後ろ倒し
        get_day = add_holyday(get_day, 1, ng_sunday)
    ElseIf send_pattern = send_other And (key = key_A Or key = key_B) Then
        '配送先パターンが各拠点で、A商品、B商品の場合
        'このパターンは固定値が算出できないため、対象外
        send_day = 0
        get_day = 0
    Else
        '上記以外のパターンは、
        '発送日が注文日+1
        '納品日が発送日から算出
        
        '発送日が土日祝だった場合、後ろ倒し
        send_day = add_holyday(rec_day + 1, 1, ng_all)
        
        '配送にかかる期間から、納期を算出
        get_day = send_day + send_time
        
        '納期が日祝だった場合、後ろ倒し
        get_day = add_holyday(get_day, 1, ng_sunday_holiday)
    End If
    
    createDay.send_day = send_day
    createDay.get_day = get_day
    
End Function

Function add_holyday(target_day, mark, ng_day)
    '日付に対して、商品によって土日祝日の回避を行う
    'markが+1の場合、翌日方向に、-1の場合、前日方向に移動する

    '「祝日表」シート保護解除
    'ThisWorkbook.Worksheets(name_sheet_holyday).Unprotect
                
    '祝日のリストを取得
    Set wsHoly = ThisWorkbook.Worksheets(name_sheet_holyday)
    With wsHoly
        last_row = .Cells(1, 1).SpecialCells(xlLastCell).row
        data_holy = .Range(.Cells(1, clm_holyday), .Cells(last_row, clm_holyday)).Value
    End With
            
    Select Case ng_day
        '日曜のみ回避するパターン
        Case ng_sunday
RET_DAY_SUN:
            week = Weekday(target_day, vbSunday)
            If week = 1 Then
                target_day = target_day + mark
                GoTo RET_DAY_SUN
            End If
            
            add_holyday = target_day
            
        '日祝を回避するパターン
        Case ng_sunday_holiday
RET_DAY_SUNHOL:
            week = Weekday(target_day, vbSunday)
            If week = 1 Then
                target_day = target_day + mark
                GoTo RET_DAY_SUNHOL
            End If
            For Each holyday In data_holy
                If target_day = holyday Then
                    target_day = target_day + mark
                    GoTo RET_DAY_SUNHOL
                End If
            Next holyday
            
            add_holyday = target_day
            
        '土日祝全て回避するパターン
        Case ng_all
RET_DAY_ALL:
            week = Weekday(target_day, vbSunday)
            If week = 7 Or week = 1 Then
                target_day = target_day + mark
                GoTo RET_DAY_ALL
            End If
            For Each holyday In data_holy
                If target_day = holyday Then
                    target_day = target_day + mark
                    GoTo RET_DAY_ALL
                End If
            Next holyday
            
            add_holyday = target_day
    End Select
    
    '「祝日表」シート保護
    'ThisWorkbook.Worksheets(name_sheet_holyday).Protect
    
End Function

Function get_send_time(target_address)
    
    '「配送先都道府県」シート保護解除
    'ThisWorkbook.Worksheets(name_sheet_area).Unprotect
    
    get_send_time = 0
    
    '都道府県のリストを取得
    Set wsArea = ThisWorkbook.Worksheets(name_sheet_area)
    With wsArea
        With .Cells(1, 1).SpecialCells(xlLastCell)
            last_row = .row
            last_clm = .Column
        End With
        data_area = .Range(.Cells(1, 1), .Cells(last_row, last_clm)).Value
    End With
    
    For idx_row = 1 To UBound(data_area, 1)
        name_area = data_area(idx_row, clm_area_name)
        If InStr(target_address, name_area) > 0 Then
            get_send_time = data_area(idx_row, clm_area_day)
            Exit For
        End If
    Next idx_row
    
    If get_send_time = 0 Then
        MsgBox "都道府県名の取得に失敗しました。" & vbCrLf & target_address
        End
    End If
        
    '「配送先都道府県」シート保護
    'ThisWorkbook.Worksheets(name_sheet_area).Protect
        
End Function