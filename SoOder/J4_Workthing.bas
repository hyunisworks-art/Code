
'このモジュールは、Excelのブックやシートに対する共通処理をまとめたものですが、
'ほかのモジュールから呼び出されている形跡がないため、将来的に不要になる可能性があります。（2026-01 西村記載）

'////////////////////////////////////////////////////////////////////////////////////
'Excelのブックやシートに対する処理
'////////////////////////////////////////////////////////////////////////////////////
'GetWsData
'GetWorkbook
'GetWorksheet
'GetLastRow
'GetLastColumn
'GetClmByName
'GetNextClmByName
'WbClose
'AddUnion
'WbSaveAs
'DeleteHiddenWs
'GetSheetName
'////////////////////////////////////////////////////////////////////////////////////

Function GetWsData(ByVal filePath As String, _
                    Optional ByVal sheetName As String = "", _
                    Optional ByVal flgResume As Boolean = False) As Variant
    '********************************************************************************
    '指定されたシートの内容を二次元配列にして返却する
    '********************************************************************************
    '   filepath   :取得したいファイルのフルパス
    '   sheetName  :取得したいシート名。指定されていない場合は、1シート目を対象とする。
    '   flgResume  :Trueが指定された場合、指定されたファイルを閉じずにそのままにしておく
    '********************************************************************************
    Dim wb As Workbook
    Set wb = GetWorkbook(filePath)
    Dim ws As Worksheet
    If sheetName = "" Then
        Set ws = wb.Worksheets(1)
    Else
        Set ws = wb.Worksheets(sheetName)
    End If
    
    Dim lastRow As Long
    lastRow = GetLastRow(ws)
    Dim lastClm As Long
    lastClm = GetLastColumn(ws)
    With ws
        GetWsData = .Range(.Cells(1, 1), .Cells(lastRow, lastClm)).Value
    End With
    
    If flgResume = False Then
        Application.DisplayAlerts = False
        wb.Close
        Application.DisplayAlerts = True
    End If
End Function

Function GetWorkbook(ByVal filePath As String, _
                        Optional ByVal readOnlyFlg As Boolean = True) As Workbook
    '********************************************************************************
    '指定されたExcelファイルを開き、開いたブックを返却する
    'すでに同名ファイルが開かれていた場合は、同名ファイルを開いたブックとして返却する
    '********************************************************************************
    '   filepath       :開きたいファイルのフルパス
    '   readOnlyFlg   :読み取り専用フラグ
    '********************************************************************************
    Dim wb As Workbook
    Dim books As Variant
    For Each books In Workbooks
        If books.Name = Dir(filePath, vbNormal) Then
            Set wb = books
            Exit For
        End If
    Next books
    
    If wb Is Nothing Then
        Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=readOnlyFlg)
    End If
    Set GetWorkbook = wb
End Function

Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    '********************************************************************************
    '指定されたブックから、指定されたシートを、ワークシート型で返却する
    '見つからない場合は、Nothingを返却する
    '********************************************************************************
    '   wb          :対象のブック
    '   sheetName  :対象のシート名
    '********************************************************************************
    Dim ws As Worksheet
    Dim wbSheet As Variant
    For Each wbSheet In wb.Worksheets
        If wbSheet.Name = sheetName Then
            Set ws = wbSheet
            Exit For
        End If
    Next wbSheet
    
    If ws Is Nothing Then
        Set GetWorksheet = Nothing
    Else
        Set GetWorksheet = ws
    End If
End Function

Function GetLastRow(ByVal ws As Worksheet, Optional ByVal clm As Long = 0) As Long
    '********************************************************************************
    '最終行を返却する
    '列を指定すると、その列の最終行を返却する
    '********************************************************************************
    '   ws  :対象のシート
    '   clm :最終行を取得したい列
    '********************************************************************************
    With ws
        Dim tmpLast As Long
        If clm = 0 Then
            tmpLast = .Cells(1, 1).SpecialCells(xlLastCell).row
        Else
            tmpLast = .Cells(.Rows.count, clm).End(xlUp).row
        End If
        
        GetLastRow = tmpLast
        
        If tmpLast <> .Rows.count Then
            Dim idxRow As Long
            For idxRow = tmpLast + 1 To .Rows.count
                If .Rows(idxRow).Hidden = False Then
                    GetLastRow = idxRow - 1
                    Exit For
                End If
            Next idxRow
        End If
    End With
End Function

Function GetLastColumn(ByVal ws As Worksheet, Optional ByVal row As Long = 0) As Long
    '********************************************************************************
    '最終列を返却する
    '行を指定すると、その行の最終列を返却する
    '********************************************************************************
    '   ws  :対象のシート
    '   row :最終列を取得したい行
    '********************************************************************************
    With ws
        Dim tmpLast As Long
        If row = 0 Then
            tmpLast = .Cells(1, 1).SpecialCells(xlLastCell).Column
        Else
            tmpLast = .Cells(row, .Columns.count).End(xlToLeft).Column
        End If
        
        GetLastColumn = tmpLast
        
        If tmpLast <> .Columns.count Then
            Dim idxClm As Long
            For idxClm = tmpLast + 1 To .Columns.count
                If .Columns(idxClm).Hidden = False Then
                    GetLastColumn = idxClm - 1
                    Exit For
                End If
            Next idxClm
        End If
    End With
End Function

Function GetClmByName(ByVal clmName As String, _
                        ByRef dataFunction As Variant, _
                        Optional ByVal targetRow As Long = 0) As Long
    '********************************************************************************
    '2次元配列から指定された名前を検索して、最初に発見した列を返却する
    '行を指定すると、他の行を無視して単一の行から検索する
    '発見できなかった場合、0を返却する
    '********************************************************************************
    '   clmName        :探したい名前
    '   dataFunction   :対象の二次元配列
    '   targetRow      :指定した場合、この行に固定して探す
    '********************************************************************************
    Dim nowClm As Long
    If targetRow > 0 Then
        For nowClm = LBound(dataFunction, 2) To UBound(dataFunction, 2)
            If dataFunction(targetRow, nowClm) = clmName Then
                GetClmByName = nowClm
                Exit Function
            End If
        Next nowClm
    Else
        Dim nowRow As Long
        For nowRow = LBound(dataFunction, 1) To UBound(dataFunction, 1)
            For nowClm = LBound(dataFunction, 2) To UBound(dataFunction, 2)
                If dataFunction(nowRow, nowClm) = clmName Then
                    GetClmByName = nowClm
                    Exit Function
                End If
            Next nowClm
        Next nowRow
    End If
    GetClmByName = 0
End Function

Function GetNextClmByName(ByVal clmName As String, _
                            ByRef dataFunction As Variant, _
                            Optional ByVal targetRow As Long = 0) As Long
    '********************************************************************************
    '2次元配列から指定された名前を検索して、その次に出現する空欄以外の列を返却する
    '発見できなかった場合、0を返却する
    '********************************************************************************
    '   clmName        :探したい名前
    '   dataFunction   :対象の二次元配列
    '   targetRow      :指定した場合、この行に固定して探す
    '********************************************************************************
    Dim nowClm As Long
    nowClm = GetClmByName(clmName, dataFunction, targetRow)
    If nowClm <> 0 Then
        Dim nowRow As Long
        Dim nextClm As Long
        For nowRow = LBound(dataFunction, 1) To UBound(dataFunction, 1)
            If dataFunction(nowRow, nowClm) = clmName Then
                For nextClm = nowClm + 1 To UBound(dataFunction, 2)
                    If dataFunction(nowRow, nextClm) <> "" Then
                        GetNextClmByName = nextClm
                        Exit Function
                    End If
                Next nextClm
            End If
        Next nowRow
    End If
    GetNextClmByName = 0
End Function

Sub WbClose(ByVal wb As Workbook)
    '********************************************************************************
    '指定されたブックを閉じる
    '********************************************************************************
    '   wb  :閉じたいブック
    '********************************************************************************
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Close
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Function AddUnion(ByRef uni As Range, ByRef rng As Range) As Range
    '********************************************************************************
    '範囲uniに、範囲rngを合算する
    '********************************************************************************
    '   uni :合算後の範囲
    '   rng :追加したい範囲
    '********************************************************************************
    If uni Is Nothing Then
        Set uni = rng
    Else
        Set uni = Union(uni, rng)
    End If
    Set AddUnion = uni
End Function

Sub WbSaveAs(ByVal wb As Workbook, ByVal path As String)
    '********************************************************************************
    '指定されたブックを別名保存する
    '********************************************************************************
    '   wb  :保存したいブック
    '   path:保存するブックのフルパス
    '********************************************************************************
    Application.DisplayAlerts = False
    Call wb.SaveAs(path)
    Application.DisplayAlerts = True
End Sub

Sub DeleteHiddenWs(ByVal wb As Workbook)
    '********************************************************************************
    '指定されたブックから、非表示シートをすべて削除する
    '********************************************************************************
    '   wb  :非表示シートを削除したいブック
    '********************************************************************************
    With wb
        Application.DisplayAlerts = False
        Dim ws As Worksheet
        For Each ws In .Worksheets
            If ws.Visible = xlSheetHidden Or _
                ws.Visible = xlSheetVeryHidden Then
                ws.Delete
            End If
        Next ws
        Application.DisplayAlerts = True
    End With
End Sub

Function GetSheetName(wb As Workbook) As String()
    '********************************************************************************
    '指定したブックのシート名を1次元配列で取得する
    '********************************************************************************
    '   wb  :シート名を取得したいブック
    '********************************************************************************
    Dim wsCount As Long
    wsCount = 0
    Dim ws As Worksheet
    Dim arrWsName() As String
    '非表示でないシート名を配列に格納
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            wsCount = wsCount + 1
            ReDim Preserve arrWsName(1 To wsCount)
            arrWsName(wsCount) = ws.Name
        End If
    Next ws
    GetSheetName = arrWsName
End Function
