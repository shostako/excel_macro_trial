Attribute VB_Name = "転記_成形データ"
Sub 成形データ転記マクロ()
    ' 各成形シートから集計表シートへのデータ転記を行う
    Call 成形データ転記("成形1", "_成形1A", "F")
    Call 成形データ転記("成形2", "_成形2A", "H")
    Call 成形データ転記("成形3", "_成形3A", "J")
    Call 成形データ転記("成形4", "_成形4A", "N")
    'Call 成形データ転記("成形5", "_成形5A", "N")
End Sub

Sub 成形データ転記(成形シート名 As String, テーブル名 As String, 出力列 As String)
    Dim ws成形 As Worksheet
    Dim ws集計表 As Worksheet
    Dim 成形範囲 As Range
    Dim 出力行 As Long
    Dim 日付 As Date
    Dim i As Long
    
    ' シートの設定
    Set ws成形 = ThisWorkbook.Sheets(成形シート名)
    Set ws集計表 = ThisWorkbook.Sheets("集計表")
    
    ' 集計表の日付を取得
    日付 = ws集計表.Range("A1").Value
    
    ' 成形シートのデータ範囲を設定
    Set 成形範囲 = ws成形.Range("I3:V34")
    
    ' 日付が合致する行を探す
    For i = 1 To 成形範囲.Rows.Count
        If ws成形.Cells(i + 2, 9).Value = 日付 Then
            出力行 = i + 2
            Exit For
        End If
    Next i
    
    ' 日付が見つかった場合、指定された列にデータを転記
    If 出力行 > 0 Then
        ws集計表.Range(出力列 & "4").Value = ws成形.Cells(出力行, 10).Value ' J列
        ws集計表.Range(出力列 & "5").Value = ws成形.Cells(出力行, 11).Value ' K列
        ws集計表.Range(出力列 & "6").Value = ws成形.Cells(出力行, 12).Value ' L列
        ws集計表.Range(出力列 & "7").Value = ws成形.Cells(出力行, 13).Value ' M列
        ws集計表.Range(出力列 & "8").Value = ws成形.Cells(出力行, 14).Value ' N列
        'ws集計表.Range(出力列 & "9").Value = ws成形.Cells(出力行, 15).Value ' O列
        ws集計表.Range(出力列 & "10").Value = ws成形.Cells(出力行, 16).Value ' P列
        ws集計表.Range(出力列 & "12").Value = ws成形.Cells(出力行, 17).Value ' Q列
        ws集計表.Range(出力列 & "13").Value = ws成形.Cells(出力行, 18).Value ' R列
        ws集計表.Range(出力列 & "14").Value = ws成形.Cells(出力行, 19).Value ' S列
        ws集計表.Range(出力列 & "15").Value = ws成形.Cells(出力行, 20).Value ' T列
        ws集計表.Range(出力列 & "16").Value = ws成形.Cells(出力行, 21).Value ' U列
        
    End If
End Sub

