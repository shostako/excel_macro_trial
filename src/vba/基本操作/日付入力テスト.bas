Sub TestMacro()
    ' エラーハンドリングを設定
    On Error GoTo ErrorHandler
    
    ' A1セルに今日の日付を入力
    Range("A1").Value = Date
    
    ' A2セルに「Hello Excel!」を入力
    Range("A2").Value = "Hello Excel!"
    
    ' 処理完了メッセージ
    MsgBox "処理が完了しました。" & vbCrLf & _
           "A1: " & Range("A1").Value & vbCrLf & _
           "A2: " & Range("A2").Value, _
           vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    ' エラーが発生した場合の処理
    MsgBox "エラーが発生しました: " & Err.Description, _
           vbCritical, "エラー"
End Sub