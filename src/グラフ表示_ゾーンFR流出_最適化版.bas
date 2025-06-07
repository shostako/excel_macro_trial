Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Option Explicit

Sub グラフ表示_ゾーンFR流出_最適化版()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 作成日: 2025/06/07
    ' 最適化版: 処理速度向上とステータスバー表示改善
    
    Dim ws As Worksheet
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String ' E3: 発生
    Dim discovery2Value As String ' E4: 発見2
    Dim arrDiscovery2 As Variant
    Dim isProcessing As Boolean    ' 「発生」が「加工」工程判定用
    Dim isMould As Boolean         ' 「発生」が「モール」工程判定用
    Dim isDiscovery2Empty As Boolean ' 発見2の値が空か判定用
    Dim commentText As String      ' D6に設定するコメント用
    
    ' エラー処理を設定
    On Error GoTo ErrorHandler
    
    ' 画面更新を停止して処理速度を向上（最重要）
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "初期設定中..."
    
    ' ワークシートの取得
    Set ws = ThisWorkbook.Worksheets("ゾーンFrRr流出")
    If ws Is Nothing Then
        MsgBox "指定されたワークシート 'ゾーンFrRr流出' が見つかりません。", vbExclamation
        GoTo Cleanup
    End If
    
    ' ピボットテーブルの取得
    Application.StatusBar = "ピボットテーブルを取得中..."
    Set pt1 = ws.PivotTables("ピボットテーブル31") ' アルヴェル Fr
    Set pt2 = ws.PivotTables("ピボットテーブル32") ' アルヴェル Rr
    Set pt3 = ws.PivotTables("ピボットテーブル33") ' ノアヴォク Fr
    Set pt4 = ws.PivotTables("ピボットテーブル34") ' ノアヴォク Rr
    Set pt5 = ws.PivotTables("ピボットテーブル35") ' モード抽出用
    
    ' ピボットテーブル取得エラーチェック
    Dim missingPT As String
    If pt1 Is Nothing Then missingPT = missingPT & "'ピボットテーブル31' "
    If pt2 Is Nothing Then missingPT = missingPT & "'ピボットテーブル32' "
    If pt3 Is Nothing Then missingPT = missingPT & "'ピボットテーブル33' "
    If pt4 Is Nothing Then missingPT = missingPT & "'ピボットテーブル34' "
    If pt5 Is Nothing Then missingPT = missingPT & "'ピボットテーブル35' "
    
    If Len(missingPT) > 0 Then
        MsgBox "指定されたピボットテーブルが見つかりません: " & vbCrLf & Trim(missingPT) & vbCrLf & _
               "シート名: '" & ws.Name & "' を確認してください。", vbExclamation
        GoTo Cleanup
    End If
    
    ' 日付範囲の取得（セルE1〜E2）
    Application.StatusBar = "パラメータを読み込み中..."
    If IsDate(ws.Range("E1").Value) And IsDate(ws.Range("E2").Value) Then
        dtStart = ws.Range("E1").Value
        dtEnd = ws.Range("E2").Value
    Else
        MsgBox "日付範囲が正しく設定されていません。セルE1とE2を確認してください。", vbExclamation
        GoTo Cleanup
    End If
    
    ' 発生値と発見2値の取得（セルE3、E4）
    occurrenceValue = Trim(CStr(ws.Range("E3").Value))
    discovery2Value = Trim(CStr(ws.Range("E4").Value))
    
    ' 発生値のエラーチェック
    If occurrenceValue = "" Then
        MsgBox "発生の値が設定されていません。セルE3を確認してください。", vbExclamation
        GoTo Cleanup
    End If
    
    ' 発見2値が空かどうかを判定
    isDiscovery2Empty = (discovery2Value = "")
    
    ' 「発生」が「加工」かどうかを判定
    isProcessing = (occurrenceValue = "加工")
    
    ' 「発生」が「モール」かどうかを判定
    isMould = (occurrenceValue = "モール")
    
    ' 発見2値をカンマ区切りで配列に分割
    If Not isDiscovery2Empty Then
        arrDiscovery2 = Split(discovery2Value, ",")
        Dim i As Long
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            arrDiscovery2(i) = Trim(arrDiscovery2(i))
        Next i
    End If
    
    ' 全ピボットテーブルの手動更新モードを一括設定（重要な最適化）
    Application.StatusBar = "ピボットテーブルの更新準備中..."
    pt1.ManualUpdate = True
    pt2.ManualUpdate = True
    pt3.ManualUpdate = True
    pt4.ManualUpdate = True
    pt5.ManualUpdate = True
    
    ' モード2フィルタをリセット（全て表示に戻す）
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(Array(pt1, pt2, pt3, pt4, pt5))
    
    ' 各ピボットテーブルのフィルタ設定（進捗表示付き）
    Application.StatusBar = "アルヴェル Fr ピボットテーブルを設定中..."
    Call OptimizedFilterPivotTable(pt1, dtStart, dtEnd, "アルヴェル", "Fr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "アルヴェル Rr ピボットテーブルを設定中..."
    Call OptimizedFilterPivotTable(pt2, dtStart, dtEnd, "アルヴェル", "Rr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Fr ピボットテーブルを設定中..."
    Call OptimizedFilterPivotTable(pt3, dtStart, dtEnd, "ノアヴォク", "Fr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Rr ピボットテーブルを設定中..."
    Call OptimizedFilterPivotTable(pt4, dtStart, dtEnd, "ノアヴォク", "Rr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call OptimizedFilterPivotTableForMode(pt5, dtStart, dtEnd, occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    ' 全ピボットテーブルを一括更新（最も重要な最適化ポイント）
    Application.StatusBar = "ピボットテーブルを更新中..."
    pt1.ManualUpdate = False
    pt2.ManualUpdate = False
    pt3.ManualUpdate = False
    pt4.ManualUpdate = False
    pt5.ManualUpdate = False
    
    ' RefreshTableは必要最小限に
    pt1.RefreshTable
    pt2.RefreshTable
    pt3.RefreshTable
    pt4.RefreshTable
    pt5.RefreshTable
    
    ' グラフ表示設定
    Application.StatusBar = "グラフ表示を設定中..."
    
    Dim showGraph1 As Boolean, showGraph2 As Boolean
    Dim showGraph3 As Boolean, showGraph4 As Boolean
    Dim startDateStr As String, endDateStr As String
    
    If isProcessing Then
        ' 「発生」が「加工」の場合
        showGraph1 = False
        showGraph2 = False
        showGraph3 = False
        showGraph4 = False
        commentText = "発生が「加工」のため、グラフは表示されません。"
    ElseIf isMould Then
        ' 「発生」が「モール」の場合
        showGraph1 = True  ' グラフ1: 表示
        showGraph2 = True  ' グラフ2: 表示
        showGraph3 = False ' グラフ3: 非表示
        showGraph4 = False ' グラフ4: 非表示
        
        ' 日付を M/D 形式に変換
        startDateStr = Format(dtStart, "m/d")
        endDateStr = Format(dtEnd, "m/d")
        commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    Else
        ' 「発生」が「加工」でも「モール」でもない場合 (その他の場合)
        ' 全てのグラフを表示
        showGraph1 = True
        showGraph2 = True
        showGraph3 = True
        showGraph4 = True
        
        ' 日付を M/D 形式に変換
        startDateStr = Format(dtStart, "m/d")
        endDateStr = Format(dtEnd, "m/d")
        commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    End If
    
    ' グラフ表示/非表示の一括適用
    Call SetChartVisibilityBatch(ws, Array("グラフ1", "グラフ2", "グラフ3", "グラフ4"), _
                                     Array(showGraph1, showGraph2, showGraph3, showGraph4))
    
    ' グラフ軸の動的調整
    Application.StatusBar = "グラフ軸を調整中..."
    
    Dim maxVal1 As Double, maxVal2 As Double
    Dim maxVal3 As Double, maxVal4 As Double
    Dim overallMax As Double
    Dim axisMax As Double
    Dim tickInterval As Double
    
    ' 各ピボットテーブルから最大値を取得
    maxVal1 = GetPivotTableMaxValue(pt1)
    maxVal2 = GetPivotTableMaxValue(pt2)
    maxVal3 = GetPivotTableMaxValue(pt3)
    maxVal4 = GetPivotTableMaxValue(pt4)
    
    ' 全体の最大値を決定
    overallMax = Application.WorksheetFunction.Max(maxVal1, maxVal2, maxVal3, maxVal4)
    
    ' 良い感じの軸最大値を計算
    axisMax = GetNiceMaxValueV3(overallMax)
    
    ' 適切な目盛り間隔を計算
    tickInterval = GetNiceTickInterval(axisMax)
    
    ' 各グラフに軸設定を適用（表示されているグラフのみ）
    If showGraph1 Then Call SetChartAxisSettings(ws, "グラフ1", axisMax, tickInterval)
    If showGraph2 Then Call SetChartAxisSettings(ws, "グラフ2", axisMax, tickInterval)
    If showGraph3 Then Call SetChartAxisSettings(ws, "グラフ3", axisMax, tickInterval)
    If showGraph4 Then Call SetChartAxisSettings(ws, "グラフ4", axisMax, tickInterval)
    
    ' D6にコメントを設定
    Application.StatusBar = "コメントを設定中..."
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの項目取得と入力規則設定
    Application.StatusBar = "モードフィールドを設定中..."
    Call SetModeFieldValidation(ws)
    
Cleanup:
    Application.StatusBar = False ' ステータスバーをクリア
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Set ws = Nothing
    Set pt1 = Nothing
    Set pt2 = Nothing
    Set pt3 = Nothing
    Set pt4 = Nothing
    Set pt5 = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "マクロエラー"
    Resume Cleanup
End Sub

Private Sub ResetMode2Filters(ByVal pivotTables As Variant)
    ' 複数のピボットテーブルのモード2フィルタを一括リセット
    Dim pt As PivotTable
    Dim i As Long
    
    On Error Resume Next
    For i = LBound(pivotTables) To UBound(pivotTables)
        Set pt = pivotTables(i)
        If Not pt Is Nothing Then
            With pt.PivotFields("モード2")
                .ClearAllFilters
                .CurrentPage = "(すべて)"
            End With
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub OptimizedFilterPivotTable(ByVal pt As PivotTable, _
                                     ByVal startDate As Date, _
                                     ByVal endDate As Date, _
                                     ByVal alNoahFilter As String, _
                                     ByVal frRrFilter As String, _
                                     ByVal occurrenceFilter As String, _
                                     ByVal discovery2Arr As Variant, _
                                     ByVal isDiscovery2Empty As Boolean)
    ' 最適化されたピボットテーブルフィルタリング
    
    Dim pi As PivotItem
    Dim d As Date
    Dim i As Long
    
    On Error Resume Next
    
    ' 日付フィールドの高速フィルタリング
    With pt.PivotFields("日付")
        .ClearAllFilters
        ' 一括で可視性を設定（ループ内でのUI更新を避ける）
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' ページフィールドの設定（これらは高速）
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールドの最適化されたフィルタリング
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            ' 最初に全て非表示にする
            For Each pi In .PivotItems
                pi.Visible = False
            Next pi
            
            ' 必要なアイテムのみ表示（重複チェックを排除）
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            
            ' 配列の要素を辞書に登録
            For i = LBound(discovery2Arr) To UBound(discovery2Arr)
                dict(discovery2Arr(i)) = True
            Next i
            
            ' 辞書に存在するアイテムのみ表示
            For Each pi In .PivotItems
                If dict.Exists(pi.Name) Then
                    pi.Visible = True
                End If
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub OptimizedFilterPivotTableForMode(ByVal pt As PivotTable, _
                                            ByVal startDate As Date, _
                                            ByVal endDate As Date, _
                                            ByVal occurrenceFilter As String, _
                                            ByVal discovery2Arr As Variant, _
                                            ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用の最適化されたフィルタリング
    
    Dim pi As PivotItem
    Dim d As Date
    Dim i As Long
    
    On Error Resume Next
    
    ' 日付フィールドの高速フィルタリング
    With pt.PivotFields("日付")
        .ClearAllFilters
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' アル/ノア・Fr/Rr：全表示（ClearAllFiltersで十分）
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド（共通処理を利用）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = False
            Next pi
            
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            
            For i = LBound(discovery2Arr) To UBound(discovery2Arr)
                dict(discovery2Arr(i)) = True
            Next i
            
            For Each pi In .PivotItems
                If dict.Exists(pi.Name) Then
                    pi.Visible = True
                End If
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByVal chartNames As Variant, ByVal visibilities As Variant)
    ' 複数のグラフの表示/非表示を一括設定
    Dim i As Long
    Dim chObj As ChartObject
    
    On Error Resume Next
    For i = LBound(chartNames) To UBound(chartNames)
        Set chObj = ws.ChartObjects(chartNames(i))
        If Not chObj Is Nothing Then
            chObj.Visible = visibilities(i)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValue(ByVal pt As PivotTable) As Double
    ' ピボットテーブルのデータ範囲から最大値を取得（最適化版）
    Dim maxVal As Double
    Dim dataRange As Range
    
    On Error Resume Next
    
    ' データ範囲を取得（値エリア）
    Set dataRange = pt.DataBodyRange
    
    If dataRange Is Nothing Then
        GetPivotTableMaxValue = 0
        Exit Function
    End If
    
    ' 配列を使用して高速化
    Dim arr As Variant
    arr = dataRange.Value
    
    If IsArray(arr) Then
        Dim r As Long, c As Long
        maxVal = 0
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If IsNumeric(arr(r, c)) And arr(r, c) > maxVal Then
                    maxVal = arr(r, c)
                End If
            Next c
        Next r
    Else
        ' 単一セルの場合
        If IsNumeric(dataRange.Value) Then
            maxVal = dataRange.Value
        Else
            maxVal = 0
        End If
    End If
    
    GetPivotTableMaxValue = maxVal
    On Error GoTo 0
End Function

Private Function GetNiceMaxValueV3(ByVal maxValue As Double) As Double
    ' データの最大値から「良い感じの」軸の最大値を計算（改良版v3）
    ' 最大値の110%～120%で最も近い切りの良い数字を選ぶ
    
    Dim minTarget As Double, maxTarget As Double
    Dim candidates As Variant
    Dim i As Long
    Dim bestValue As Double
    Dim bestDiff As Double
    Dim currentDiff As Double
    
    If maxValue <= 0 Then
        GetNiceMaxValueV3 = 10
        Exit Function
    End If
    
    ' 目標範囲：最大値の110%～120%
    minTarget = maxValue * 1.1
    maxTarget = maxValue * 1.2
    
    ' 切りの良い数字の候補を生成
    candidates = GenerateNiceCandidates(minTarget, maxTarget)
    
    ' 最も近い候補を選択
    bestValue = candidates(0)
    bestDiff = Abs(candidates(0) - maxValue)
    
    For i = 1 To UBound(candidates)
        currentDiff = Abs(candidates(i) - maxValue)
        If currentDiff < bestDiff And candidates(i) >= minTarget Then
            bestValue = candidates(i)
            bestDiff = currentDiff
        End If
    Next i
    
    ' 最小でも最大値+1は保証
    If bestValue <= maxValue Then
        bestValue = maxValue + 1
    End If
    
    GetNiceMaxValueV3 = bestValue
End Function

Private Function GenerateNiceCandidates(ByVal minVal As Double, ByVal maxVal As Double) As Variant
    ' 指定範囲内の切りの良い数字の候補を生成（最適化版）
    Dim candidates() As Double
    Dim count As Long
    Dim value As Double
    
    ' 基本的な倍数パターン
    Dim multipliers As Variant
    multipliers = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 15, 20, 25, 30, 40, 50, 60, 70, 80, 90)
    
    ReDim candidates(0 To 100) ' 十分な容量を確保
    count = 0
    
    ' 桁数を計算
    Dim magnitude As Long
    magnitude = Int(Log(maxVal) / Log(10))
    
    ' 対象となる基数を限定
    Dim base As Long
    Dim startBase As Long
    startBase = 10 ^ (magnitude - 1)
    
    Dim i As Long
    For base = startBase To 10 ^ (magnitude + 1) Step startBase
        For i = LBound(multipliers) To UBound(multipliers)
            value = multipliers(i) * base
            If value >= minVal - 5 And value <= maxVal + 5 Then
                candidates(count) = value
                count = count + 1
                If count > UBound(candidates) Then Exit For
            End If
        Next i
        If count > 50 Then Exit For ' 十分な候補が得られたら終了
    Next base
    
    ' 配列をリサイズ
    If count > 0 Then
        ReDim Preserve candidates(0 To count - 1)
    Else
        ReDim candidates(0 To 0)
        candidates(0) = maxVal + 1
    End If
    
    GenerateNiceCandidates = candidates
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 軸の最大値に基づいて適切な目盛り間隔を計算
    ' 目標：軸に5～10本程度の目盛り線
    
    Dim targetTicks As Long
    Dim roughInterval As Double
    Dim niceInterval As Double
    
    ' 理想的な目盛り数は6～8本
    targetTicks = 6
    roughInterval = maxValue / targetTicks
    
    ' 切りの良い間隔に調整
    Select Case True
        Case roughInterval <= 1: niceInterval = 1
        Case roughInterval <= 2: niceInterval = 2
        Case roughInterval <= 5: niceInterval = 5
        Case roughInterval <= 10: niceInterval = 10
        Case roughInterval <= 20: niceInterval = 20
        Case roughInterval <= 25: niceInterval = 25
        Case roughInterval <= 50: niceInterval = 50
        Case roughInterval <= 100: niceInterval = 100
        Case roughInterval <= 200: niceInterval = 200
        Case roughInterval <= 250: niceInterval = 250
        Case roughInterval <= 500: niceInterval = 500
        Case Else
            ' 1000以上の場合
            Dim magnitude As Long
            magnitude = Int(Log(roughInterval) / Log(10))
            Dim base As Double
            base = 10 ^ magnitude
            
            Select Case True
                Case roughInterval <= 2 * base: niceInterval = 2 * base
                Case roughInterval <= 5 * base: niceInterval = 5 * base
                Case Else: niceInterval = 10 * base
            End Select
    End Select
    
    GetNiceTickInterval = niceInterval
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, ByVal chartName As String, ByVal maxValue As Double, ByVal tickInterval As Double)
    ' グラフの縦軸設定（最大値と目盛り間隔）
    Dim chObj As ChartObject
    Dim ch As Chart
    
    On Error Resume Next
    Set chObj = ws.ChartObjects(chartName)
    
    If Not chObj Is Nothing Then
        Set ch = chObj.Chart
        
        ' Y軸（縦軸）の設定
        With ch.Axes(xlValue)
            .MaximumScaleIsAuto = False
            .MaximumScale = maxValue
            .MinimumScaleIsAuto = False
            .MinimumScale = 0
            
            ' 主目盛り間隔の設定
            .MajorUnitIsAuto = False
            .MajorUnit = tickInterval
            
            ' 補助目盛りは主目盛りの半分（必要に応じて）
            .MinorUnitIsAuto = False
            .MinorUnit = tickInterval / 2
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub SetModeFieldValidation(ByVal ws As Worksheet)
    ' モードフィールドの入力規則設定（最適化版）
    Dim modeItems As Object
    Dim modeList As String
    Dim lastRow As Long
    Dim cellValue As String
    Dim excludeDict As Object
    Dim arr As Variant
    Dim i As Long
    
    ' Dictionary使って重複排除
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    ' 除外する値を辞書で設定（高速チェック用）
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.Add "A", True
    excludeDict.Add "B", True
    excludeDict.Add "C", True
    excludeDict.Add "D", True
    excludeDict.Add "E", True
    excludeDict.Add "Fr RH", True
    
    On Error Resume Next
    
    ' AG列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        ' 配列を使用して高速化
        arr = ws.Range("AG13:AG" & lastRow).Value
        
        If IsArray(arr) Then
            For i = 1 To UBound(arr, 1)
                cellValue = Trim(CStr(arr(i, 1)))
                
                If cellValue <> "" And Not excludeDict.Exists(cellValue) And Not modeItems.Exists(cellValue) Then
                    modeItems.Add cellValue, cellValue
                End If
            Next i
        End If
    End If
    
    On Error GoTo 0
    
    ' リスト文字列作成
    If modeItems.Count > 0 Then
        modeList = Join(modeItems.Keys, ",")
        
        ' T3セルに入力規則設定
        With ws.Range("T3")
            .Validation.Delete
            .Value = "" ' 古い値をクリア
            .Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=modeList
        End With
    Else
        ' モード項目が見つからない場合
        ws.Range("T3").Validation.Delete
        ws.Range("T3").Value = "モード項目なし"
    End If
End Sub