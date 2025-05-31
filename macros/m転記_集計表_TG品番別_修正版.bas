Attribute VB_Name = "m転記_集計表_TG品番別"
Option Explicit

' 転記_集計表_TG品番別
' 集計表のA1セルの日付を基に、TG品番別シートから該当データを転記
Sub 転記_集計表_TG品番別()
    ' 高速化設定（最優先）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 変数宣言
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim targetDate As Date
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long, k As Long
    Dim sourceRow As Long
    Dim totalCombinations As Long
    Dim processedCombinations As Long
    
    ' 品番接頭辞の配列
    Dim prefixList() As Variant
    prefixList = Array("RH", "LH", "合計")
    
    ' 転記元列名末尾の配列
    Dim suffixList() As Variant
    suffixList = Array("日実績", "日出来高ｻｲｸﾙ", "累計実績", "平均実績", _
                      "累計出来高ｻｲｸﾙ", "日不良実績", "累計不良数", _
                      "累計不良率", "平均不良数")
    
    ' 転記先行番号の配列（suffixListに対応）
    Dim targetRows() As Variant
    targetRows = Array(33, 34, 35, 36, 37, 39, 40, 41, 42)
    
    ' 品番に対応する転記先列の配列
    Dim targetColumns() As Variant
    targetColumns = Array(12, 14, 16)  ' L, N, P列
    
    On Error GoTo ErrorHandler
    
    ' ========== 基本設定 ==========
    Application.StatusBar = "TGデータの転記処理を開始します..."
    
    ' ========== データソース確認 ==========
    Application.StatusBar = "データソース確認中..."
    
    ' 集計表シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' 集計表のA1セルから日付取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。", vbCritical
        GoTo Cleanup
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' TG品番別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("TG品番別")
    If wsSource Is Nothing Then
        MsgBox "「TG品番別」シートが見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_TG品番別b")
    If sourceTable Is Nothing Then
        MsgBox "「TG品番別」シートに「_TG品番別b」テーブルが見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_TG品番別b」テーブルにデータがありません。", vbCritical
        GoTo Cleanup
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ========== 列インデックス取得 ==========
    Application.StatusBar = "列構造解析中..."
    
    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_TG品番別b」テーブルに「日付」列が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' ========== 該当日付の行を検索 ==========
    Application.StatusBar = "該当日付検索中..."
    sourceRow = 0
    For j = 1 To sourceData.Rows.Count
        If sourceData.Cells(j, dateColIndex).Value = targetDate Then
            sourceRow = j
            Exit For
        End If
    Next j
    
    If sourceRow = 0 Then
        MsgBox "日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    ' ========== 転記処理準備 ==========
    Application.StatusBar = "転記処理準備中..."
    
    ' 各品番と末尾の組み合わせで転記処理
    totalCombinations = (UBound(prefixList) + 1) * (UBound(suffixList) + 1)
    processedCombinations = 0
    
    ' ========== データ転記処理 ==========
    Application.StatusBar = "データ転記処理中..."
    
    For i = 0 To UBound(prefixList)
        For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1
            
            ' 列名を構築（品番接頭辞 + 末尾文字列）
            Dim columnName As String
            columnName = prefixList(i) & suffixList(k)
            
            ' 転記実行
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(columnName).Index
            
            If Err.Number = 0 Then
                ' ソース値を一旦変数に格納
                Dim sourceValue As Variant
                sourceValue = sourceData.Cells(sourceRow, colIndex).Value
                
                ' 空白チェックと転記
                If IsEmpty(sourceValue) Or sourceValue = "" Or IsNull(sourceValue) Then
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = 0
                Else
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = sourceValue
                End If
            Else
                ' 列が見つからない場合は警告（デバッグ用）
                Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' 進捗更新（10件ごと）
            If processedCombinations Mod 10 = 0 Then
                Application.StatusBar = "TGデータ転記中... (" & _
                    processedCombinations & "/" & totalCombinations & ")"
                DoEvents
            End If
        Next k
    Next i
    
    ' ========== 処理完了 ==========
    Application.StatusBar = "転記処理が完了しました"
    
    ' 正常終了（エラー時以外はメッセージ非表示）
    GoTo Cleanup
    
ErrorHandler:
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    
Cleanup:
    ' 設定を元に戻す
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' オブジェクト解放
    Set wsTarget = Nothing
    Set wsSource = Nothing
    Set sourceTable = Nothing
    Set sourceData = Nothing
End Sub