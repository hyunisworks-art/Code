Private Sub W2Pの制御03()
' W2Pデータシートのクリアのモジュールだが呼び出されている痕跡がないので
' 将来的に不要になる可能性が高い。（2026/1/19西村記載）
    Dim end_data_row As Long
    If MsgBox("リストを全て削除します。よろしいですか？", vbOKCancel) = vbOK Then
        With ThisWorkbook.Worksheets(w2pdata_sheet)
        end_data_row = .Cells(.Rows.count, 35).End(xlUp).row
        .Range(.Cells(2, 1), .Cells(end_data_row, 34)).Value = ""
        .Range(.Rows(2), .Rows(end_data_row)).Interior.ColorIndex = syokika_color
        End With
        MsgBox "完了しました。"
    End If
End Sub