Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Sub グラフ表示_ゾーンFR流出()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 最適化版 - 処理速度改善とステータスバー詳細化
    ' 作成日: 2025/06/07
    
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
    
    ' 画面更新を停止して処理速度を向上（三種の神器）
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "初期化中..."
    
    ' ワークシートの取得
    Set ws = ThisWorkbook.Worksheets("ゾーンFrRr流出")
    
    ' ピボットテーブルの取得
    Application.StatusBar = "ピボットテーブルを取得中..."
    Set pt1 = ws.PivotTables("ピボットテーブル31") ' アルヴェル Fr
    Set pt2 = ws.PivotTables("ピボットテーブル32") ' アルヴェル Rr
    Set pt3 = ws.PivotTables("ピボットテーブル33") ' ノアヴォク Fr
    Set pt4 = ws.PivotTables("ピボットテーブル34") ' ノアヴォク Rr
    Set pt5 = ws.PivotTables("ピボットテーブル35") ' モード抽出用
    
    ' 日付範囲の取得（セルE1～E2）
    dtStart = ws.Range("E1").Value
    dtEnd = ws.Range("E2").Value
    
    ' 発生値と発見2値の取得（セルE3、E4）
    occurrenceValue = Trim(CStr(ws.Range("E3").Value))
    discovery2Value = Trim(CStr(ws.Range("E4").Value))
    
    ' 発見2値が空かどうかを判定
    isDiscovery2Empty = (discovery2Value = "")
    
    ' 「発生」が「加工」かどうかを判定
    isProcessing = (occurrenceValue = "加工")
    
    ' 「発生」が「モール」かどうかを判定
    isMould = (occurrenceValue = "モール")
    
    ' 発見2値をカンマ区切りで配列に分割（Dictionaryで高速化）
    Dim discovery2Dict As Object
    Set discovery2Dict = CreateObject("Scripting.Dictionary")
    
    If Not isDiscovery2Empty Then
        arrDiscovery2 = Split(discovery2Value, ",")
        Dim i As Long
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            discovery2Dict(Trim(arrDiscovery2(i))) = True
        Next i
    End If
    
    ' 全ピボットテーブルの手動更新モードを一括設定（高速化の要）
    Application.StatusBar = "ピボットテーブルの更新モードを設定中..."
    pt1.ManualUpdate = True
    pt2.ManualUpdate = True
    pt3.ManualUpdate = True
    pt4.ManualUpdate = True
    pt5.ManualUpdate = True
    
    ' モード2フィルタをリセット（一括処理）
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(Array(pt1, pt2, pt3, pt4, pt5))
    
    ' 各ピボットテーブルのフィルタ設定（個別にステータス表示）
    Application.StatusBar = "アルヴェル Fr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt1, dtStart, dtEnd, "アルヴェル", "Fr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "アルヴェル Rr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt2, dtStart, dtEnd, "アルヴェル", "Rr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Fr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt3, dtStart, dtEnd, "ノアヴォク", "Fr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Rr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt4, dtStart, dtEnd, "ノアヴォク", "Rr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call FilterPivotTableForModeOptimized(pt5, dtStart, dtEnd, occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    ' ピボットテーブルの一括更新（ここが高速化のポイント）
    Application.StatusBar = "全ピボットテーブルを更新中..."
    pt1.ManualUpdate = False
    pt2.ManualUpdate = False
    pt3.ManualUpdate = False
    pt4.ManualUpdate = False
    pt5.ManualUpdate = False
    
    ' RefreshTableは最後に一括実行
    pt1.RefreshTable
    pt2.RefreshTable
    pt3.RefreshTable
    pt4.RefreshTable
    pt5.RefreshTable
    
    ' グラフ表示設定
    Application.StatusBar = "グラフ表示設定を適用中..."
    
    Dim showGraphs(1 To 4) As Boolean
    Dim startDateStr As String, endDateStr As String
    
    Select Case True
        Case isProcessing
            ' 「発生」が「加工」の場合
            showGraphs(1) = False
            showGraphs(2) = False
            showGraphs(3) = False
            showGraphs(4) = False
            commentText = "発生が「加工」のため、グラフは表示されません。"
            
        Case isMould
            ' 「発生」が「モール」の場合
            showGraphs(1) = True
            showGraphs(2) = True
            showGraphs(3) = False
            showGraphs(4) = False
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
            
        Case Else
            ' その他の場合
            showGraphs(1) = True
            showGraphs(2) = True
            showGraphs(3) = True
            showGraphs(4) = True
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    End Select
    
    ' グラフ表示/非表示の一括適用
    Call SetChartVisibilityBatch(ws, showGraphs)
    
    ' グラフ軸の動的調整
    Application.StatusBar = "グラフ軸を調整中..."
    
    ' 各ピボットテーブルから最大値を取得（配列処理で高速化）
    Dim maxValues(1 To 4) As Double
    maxValues(1) = GetPivotTableMaxValueFast(pt1)
    maxValues(2) = GetPivotTableMaxValueFast(pt2)
    maxValues(3) = GetPivotTableMaxValueFast(pt3)
    maxValues(4) = GetPivotTableMaxValueFast(pt4)
    
    ' 全体の最大値を決定
    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxValues)
    
    ' 良い感じの軸最大値と目盛り間隔を計算
    Dim axisMax As Double, tickInterval As Double
    axisMax = GetNiceMaxValueOptimized(overallMax)
    tickInterval = GetNiceTickInterval(axisMax)
    
    ' 各グラフに軸設定を適用（表示されているグラフのみ）
    Dim j As Long
    For j = 1 To 4
        If showGraphs(j) Then
            Call SetChartAxisSettings(ws, "グラフ" & j, axisMax, tickInterval)
        End If
    Next j
    
    ' D6にコメントを設定
    Application.StatusBar = "最終設定を適用中..."
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの項目取得と入力規則設定（配列処理で高速化）
    Dim modeItems As Object
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    ' 除外する値を事前定義
    Dim excludeItems As Object
    Set excludeItems = CreateObject("Scripting.Dictionary")
    excludeItems.Add "A", True
    excludeItems.Add "B", True
    excludeItems.Add "C", True
    excludeItems.Add "D", True
    excludeItems.Add "E", True
    excludeItems.Add "Fr RH", True
    
    ' AG列の最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        ' 配列で一括読み込み（高速化）
        Dim rngData As Variant
        rngData = ws.Range("AG13:AG" & lastRow).Value
        
        Dim rowIndex As Long
        For rowIndex = 1 To UBound(rngData, 1)
            Dim cellValue As String
            cellValue = Trim(CStr(rngData(rowIndex, 1)))
            
            If cellValue <> "" And Not excludeItems.Exists(cellValue) Then
                modeItems(cellValue) = True
            End If
        Next rowIndex
    End If
    
    ' リスト文字列作成
    If modeItems.Count > 0 Then
        Dim modeList As String
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
    
Cleanup:
    Application.StatusBar = "処理が完了しました。"
    Application.Wait Now + TimeValue("00:00:01") ' 1秒間表示
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

Private Sub FilterPivotTableOptimized(ByVal pt As PivotTable, _
                                     ByVal startDate As Date, _
                                     ByVal endDate As Date, _
                                     ByVal alNoahFilter As String, _
                                     ByVal frRrFilter As String, _
                                     ByVal occurrenceFilter As String, _
                                     ByVal discovery2Dict As Object, _
                                     ByVal isDiscovery2Empty As Boolean)
    ' 最適化されたピボットテーブルフィルタリング
    
    On Error Resume Next
    
    ' 日付フィールドのフィルタリング（効率的な処理）
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim pi As PivotItem
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                Dim d As Date
                d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' 各フィールドの設定（CurrentPageを使用）
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールドのフィルタリング（Dictionary使用で高速化）
    If Not isDiscovery2Empty Then
        With pt.PivotFields("発見2")
            .ClearAllFilters
            ' 全アイテムを一旦非表示
            For Each pi In .PivotItems
                pi.Visible = False
            Next pi
            ' Dictionary内のアイテムのみ表示
            For Each pi In .PivotItems
                If discovery2Dict.Exists(pi.Name) Then
                    pi.Visible = True
                End If
            Next pi
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub FilterPivotTableForModeOptimized(ByVal pt As PivotTable, _
                                            ByVal startDate As Date, _
                                            ByVal endDate As Date, _
                                            ByVal occurrenceFilter As String, _
                                            ByVal discovery2Dict As Object, _
                                            ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用の最適化されたフィルタリング
    
    On Error Resume Next
    
    ' 日付フィールドのフィルタリング
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim pi As PivotItem
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                Dim d As Date
                d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' アル/ノア・Fr/Rr：全表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド（Dictionary使用）
    If Not isDiscovery2Empty Then
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = False
            Next pi
            For Each pi In .PivotItems
                If discovery2Dict.Exists(pi.Name) Then
                    pi.Visible = True
                End If
            Next pi
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub ResetMode2Filters(ByVal pivotTables As Variant)
    ' モード2フィルタの一括リセット
    Dim pt As PivotTable
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To UBound(pivotTables)
        Set pt = pivotTables(i)
        With pt.PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = "(すべて)"
        End With
    Next i
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByRef showGraphs() As Boolean)
    ' グラフの表示/非表示を一括設定
    Dim i As Long
    On Error Resume Next
    For i = 1 To 4
        ws.ChartObjects("グラフ" & i).Visible = showGraphs(i)
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を配列処理で高速取得
    Dim maxVal As Double
    Dim dataRange As Range
    
    On Error Resume Next
    Set dataRange = pt.DataBodyRange
    
    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' 配列で一括読み込み
    Dim arr As Variant
    arr = dataRange.Value
    
    Dim i As Long, j As Long
    maxVal = 0
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If IsNumeric(arr(i, j)) And arr(i, j) > maxVal Then
                maxVal = arr(i, j)
            End If
        Next j
    Next i
    
    GetPivotTableMaxValueFast = maxVal
    On Error GoTo 0
End Function

Private Function GetNiceMaxValueOptimized(ByVal maxValue As Double) As Double
    ' 最適化された軸最大値計算
    
    If maxValue <= 0 Then
        GetNiceMaxValueOptimized = 10
        Exit Function
    End If
    
    ' 目標範囲：最大値の110%～120%
    Dim minTarget As Double, maxTarget As Double
    minTarget = maxValue * 1.1
    maxTarget = maxValue * 1.2
    
    ' 効率的な候補生成
    Dim magnitude As Long
    magnitude = Int(Log(minTarget) / Log(10))
    
    Dim base As Double
    base = 10 ^ magnitude
    
    ' 候補値を直接計算
    Dim candidates As Variant
    candidates = Array(base, base * 2, base * 5, base * 10)
    
    Dim i As Long
    For i = 0 To UBound(candidates)
        If candidates(i) >= minTarget And candidates(i) <= maxTarget Then
            GetNiceMaxValueOptimized = candidates(i)
            Exit Function
        End If
    Next i
    
    ' 適切な候補が見つからない場合
    GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(minTarget, base)
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 軸の最大値に基づいて適切な目盛り間隔を計算
    Dim targetTicks As Long
    Dim roughInterval As Double
    
    targetTicks = 6
    roughInterval = maxValue / targetTicks
    
    ' 効率的な間隔決定
    Select Case True
        Case roughInterval <= 1: GetNiceTickInterval = 1
        Case roughInterval <= 2: GetNiceTickInterval = 2
        Case roughInterval <= 5: GetNiceTickInterval = 5
        Case roughInterval <= 10: GetNiceTickInterval = 10
        Case roughInterval <= 20: GetNiceTickInterval = 20
        Case roughInterval <= 25: GetNiceTickInterval = 25
        Case roughInterval <= 50: GetNiceTickInterval = 50
        Case roughInterval <= 100: GetNiceTickInterval = 100
        Case roughInterval <= 200: GetNiceTickInterval = 200
        Case roughInterval <= 250: GetNiceTickInterval = 250
        Case roughInterval <= 500: GetNiceTickInterval = 500
        Case Else
            Dim magnitude As Long
            magnitude = Int(Log(roughInterval) / Log(10))
            Dim base As Double
            base = 10 ^ magnitude
            
            Select Case True
                Case roughInterval <= 2 * base: GetNiceTickInterval = 2 * base
                Case roughInterval <= 5 * base: GetNiceTickInterval = 5 * base
                Case Else: GetNiceTickInterval = 10 * base
            End Select
    End Select
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, ByVal chartName As String, ByVal maxValue As Double, ByVal tickInterval As Double)
    ' グラフの縦軸設定
    Dim chObj As ChartObject
    
    On Error Resume Next
    Set chObj = ws.ChartObjects(chartName)
    
    If Not chObj Is Nothing Then
        With chObj.Chart.Axes(xlValue)
            .MaximumScaleIsAuto = False
            .MaximumScale = maxValue
            .MinimumScaleIsAuto = False
            .MinimumScale = 0
            .MajorUnitIsAuto = False
            .MajorUnit = tickInterval
            .MinorUnitIsAuto = False
            .MinorUnit = tickInterval / 2
        End With
    End If
    
    Set chObj = Nothing
    On Error GoTo 0
End Sub