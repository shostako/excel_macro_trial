Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Option Explicit

Sub グラフ表示_ゾーンFR流出()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 最適化版: 処理速度改善と詳細な進捗表示
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

    ' 全ピボットテーブルの手動更新モードを一括設定（高速化の要）
    Application.StatusBar = "ピボットテーブル更新モードを設定中..."
    Call SetAllPivotTablesManualUpdate(Array(pt1, pt2, pt3, pt4, pt5), True)

    ' モード2フィルタを一括リセット
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2FiltersOptimized(Array(pt1, pt2, pt3, pt4, pt5))

    ' 各ピボットテーブルのフィルタ設定（詳細な進捗表示付き）
    Application.StatusBar = "アルヴェル Fr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt1, dtStart, dtEnd, "アルヴェル", "Fr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "アルヴェル Rr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt2, dtStart, dtEnd, "アルヴェル", "Rr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Fr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt3, dtStart, dtEnd, "ノアヴォク", "Fr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Rr ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pt4, dtStart, dtEnd, "ノアヴォク", "Rr", occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call FilterPivotTableForModeOptimized(pt5, dtStart, dtEnd, occurrenceValue, arrDiscovery2, isDiscovery2Empty)

    ' 全ピボットテーブルを一括更新（最も効率的）
    Application.StatusBar = "全ピボットテーブルを更新中..."
    Call SetAllPivotTablesManualUpdate(Array(pt1, pt2, pt3, pt4, pt5), False)
    Call RefreshAllPivotTables(Array(pt1, pt2, pt3, pt4, pt5))

    ' グラフ表示設定の準備
    Application.StatusBar = "グラフ表示設定を準備中..."
    Dim showGraph1 As Boolean, showGraph2 As Boolean
    Dim showGraph3 As Boolean, showGraph4 As Boolean
    Dim startDateStr As String, endDateStr As String

    ' グラフ表示ロジックの最適化（Select Case使用）
    Select Case True
        Case isProcessing
            ' 「発生」が「加工」の場合
            showGraph1 = False
            showGraph2 = False
            showGraph3 = False
            showGraph4 = False
            commentText = "発生が「加工」のため、グラフは表示されません。"
            
        Case isMould
            ' 「発生」が「モール」の場合
            showGraph1 = True
            showGraph2 = True
            showGraph3 = False
            showGraph4 = False
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
            
        Case Else
            ' その他の場合
            showGraph1 = True
            showGraph2 = True
            showGraph3 = True
            showGraph4 = True
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    End Select

    ' グラフ表示/非表示を一括設定
    Application.StatusBar = "グラフ表示を設定中..."
    Call SetChartVisibilityBatch(ws, Array("グラフ1", "グラフ2", "グラフ3", "グラフ4"), _
                                  Array(showGraph1, showGraph2, showGraph3, showGraph4))

    ' グラフ軸の動的調整（配列使用で高速化）
    Application.StatusBar = "グラフ軸を調整中..."
    Call AdjustChartAxesOptimized(ws, Array(pt1, pt2, pt3, pt4), _
                                  Array("グラフ1", "グラフ2", "グラフ3", "グラフ4"), _
                                  Array(showGraph1, showGraph2, showGraph3, showGraph4))
    
    ' D6にコメントを設定
    Application.StatusBar = "コメントを設定中..."
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの入力規則設定（最適化版）
    Application.StatusBar = "モードフィールドを設定中..."
    Call SetModeFieldValidationOptimized(ws)

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

' ヘルパー関数群（最適化版）

Private Sub SetAllPivotTablesManualUpdate(ByVal pivotTables As Variant, ByVal isManual As Boolean)
    ' 全ピボットテーブルの更新モードを一括設定
    Dim pt As Variant
    For Each pt In pivotTables
        If Not pt Is Nothing Then
            pt.ManualUpdate = isManual
        End If
    Next pt
End Sub

Private Sub RefreshAllPivotTables(ByVal pivotTables As Variant)
    ' 全ピボットテーブルを一括更新
    Dim pt As Variant
    For Each pt In pivotTables
        If Not pt Is Nothing Then
            pt.RefreshTable
        End If
    Next pt
End Sub

Private Sub ResetMode2FiltersOptimized(ByVal pivotTables As Variant)
    ' モード2フィルタを効率的にリセット
    Dim pt As Variant
    On Error Resume Next
    For Each pt In pivotTables
        If Not pt Is Nothing Then
            With pt.PivotFields("モード2")
                .ClearAllFilters
                .CurrentPage = "(すべて)"
            End With
        End If
    Next pt
    On Error GoTo 0
End Sub

Private Sub FilterPivotTableOptimized(ByVal pt As PivotTable, _
                                      ByVal startDate As Date, _
                                      ByVal endDate As Date, _
                                      ByVal alNoahFilter As String, _
                                      ByVal frRrFilter As String, _
                                      ByVal occurrenceFilter As String, _
                                      ByVal discovery2Arr As Variant, _
                                      ByVal isDiscovery2Empty As Boolean)
    ' 最適化されたピボットテーブルフィルタリング
    
    On Error Resume Next
    
    ' 日付フィールドの高速フィルタリング（Dictionary使用）
    Dim dateDict As Object
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    ' 有効な日付を事前に辞書に格納
    Dim d As Date
    For d = startDate To endDate
        dateDict(Format(d, "yyyy/mm/dd")) = True
        dateDict(Format(d, "yyyy/m/d")) = True ' 日付形式のバリエーションに対応
    Next d
    
    ' 日付フィルタリング（一括処理）
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim pi As PivotItem
        For Each pi In .PivotItems
            pi.Visible = dateDict.Exists(pi.Name)
        Next pi
    End With
    
    ' レポートフィルタの設定（CurrentPageは高速）
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィルタリング（Dictionary使用で高速化）
    If Not isDiscovery2Empty Then
        Dim discovery2Dict As Object
        Set discovery2Dict = CreateObject("Scripting.Dictionary")
        
        ' 選択する項目を辞書に格納
        Dim i As Long
        For i = LBound(discovery2Arr) To UBound(discovery2Arr)
            discovery2Dict(discovery2Arr(i)) = True
        Next i
        
        ' 一括フィルタリング
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub FilterPivotTableForModeOptimized(ByVal pt As PivotTable, _
                                             ByVal startDate As Date, _
                                             ByVal endDate As Date, _
                                             ByVal occurrenceFilter As String, _
                                             ByVal discovery2Arr As Variant, _
                                             ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用の最適化フィルタリング
    
    On Error Resume Next
    
    ' 日付フィルタリング（同様に最適化）
    Dim dateDict As Object
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    Dim d As Date
    For d = startDate To endDate
        dateDict(Format(d, "yyyy/mm/dd")) = True
        dateDict(Format(d, "yyyy/m/d")) = True
    Next d
    
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim pi As PivotItem
        For Each pi In .PivotItems
            pi.Visible = dateDict.Exists(pi.Name)
        Next pi
    End With
    
    ' アル/ノア、Fr/Rrは全て表示（フィルタクリアのみ）
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィルタ
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィルタ（同様に最適化）
    If Not isDiscovery2Empty Then
        Dim discovery2Dict As Object
        Set discovery2Dict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        For i = LBound(discovery2Arr) To UBound(discovery2Arr)
            discovery2Dict(discovery2Arr(i)) = True
        Next i
        
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByVal chartNames As Variant, ByVal visibilities As Variant)
    ' グラフの表示/非表示を一括設定
    Dim i As Long
    On Error Resume Next
    For i = LBound(chartNames) To UBound(chartNames)
        ws.ChartObjects(chartNames(i)).Visible = visibilities(i)
    Next i
    On Error GoTo 0
End Sub

Private Sub AdjustChartAxesOptimized(ByVal ws As Worksheet, _
                                     ByVal pivotTables As Variant, _
                                     ByVal chartNames As Variant, _
                                     ByVal visibilities As Variant)
    ' グラフ軸の動的調整（最適化版）
    
    ' 各ピボットテーブルから最大値を高速取得
    Dim maxValues() As Double
    ReDim maxValues(LBound(pivotTables) To UBound(pivotTables))
    Dim i As Long
    
    For i = LBound(pivotTables) To UBound(pivotTables)
        maxValues(i) = GetPivotTableMaxValueFast(pivotTables(i))
    Next i
    
    ' 全体の最大値
    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxValues)
    
    ' 軸設定の計算
    Dim axisMax As Double
    Dim tickInterval As Double
    axisMax = GetNiceMaxValue(overallMax)
    tickInterval = GetNiceTickInterval(axisMax)
    
    ' 表示されているグラフにのみ軸設定を適用
    On Error Resume Next
    For i = LBound(chartNames) To UBound(chartNames)
        If visibilities(i) Then
            With ws.ChartObjects(chartNames(i)).Chart.Axes(xlValue)
                .MaximumScaleIsAuto = False
                .MaximumScale = axisMax
                .MinimumScaleIsAuto = False
                .MinimumScale = 0
                .MajorUnitIsAuto = False
                .MajorUnit = tickInterval
                .MinorUnitIsAuto = False
                .MinorUnit = tickInterval / 2
            End With
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を高速取得（配列使用）
    On Error Resume Next
    
    If pt.DataBodyRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' データを配列に読み込み
    Dim dataArr As Variant
    dataArr = pt.DataBodyRange.Value
    
    If Not IsArray(dataArr) Then
        GetPivotTableMaxValueFast = CDbl(dataArr)
        Exit Function
    End If
    
    ' 配列内の最大値を検索
    Dim maxVal As Double
    Dim i As Long, j As Long
    maxVal = 0
    
    For i = 1 To UBound(dataArr, 1)
        For j = 1 To UBound(dataArr, 2)
            If IsNumeric(dataArr(i, j)) And dataArr(i, j) > maxVal Then
                maxVal = dataArr(i, j)
            End If
        Next j
    Next i
    
    GetPivotTableMaxValueFast = maxVal
    On Error GoTo 0
End Function

Private Function GetNiceMaxValue(ByVal maxValue As Double) As Double
    ' 良い感じの軸最大値を計算（簡略化版）
    If maxValue <= 0 Then
        GetNiceMaxValue = 10
        Exit Function
    End If
    
    Dim targetValue As Double
    targetValue = maxValue * 1.15 ' 最大値の115%
    
    ' 切りの良い数字に調整
    Dim magnitude As Long
    magnitude = Int(Log(targetValue) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude
    
    If targetValue <= base Then
        GetNiceMaxValue = base
    ElseIf targetValue <= 2 * base Then
        GetNiceMaxValue = 2 * base
    ElseIf targetValue <= 5 * base Then
        GetNiceMaxValue = 5 * base
    Else
        GetNiceMaxValue = 10 * base
    End If
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 適切な目盛り間隔を計算
    Dim targetTicks As Long
    targetTicks = 6
    Dim roughInterval As Double
    roughInterval = maxValue / targetTicks
    
    ' 切りの良い間隔に調整（配列使用で高速化）
    Dim intervals As Variant
    intervals = Array(1, 2, 5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000, 5000, 10000)
    
    Dim i As Long
    For i = 0 To UBound(intervals)
        If roughInterval <= intervals(i) Then
            GetNiceTickInterval = intervals(i)
            Exit Function
        End If
    Next i
    
    ' より大きな値の場合
    GetNiceTickInterval = intervals(UBound(intervals)) * Application.RoundUp(roughInterval / intervals(UBound(intervals)), 0)
End Function

Private Sub SetModeFieldValidationOptimized(ByVal ws As Worksheet)
    ' モードフィールドの入力規則を最適化設定
    
    ' AG列のデータを配列で一括取得（高速化）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow < 13 Then
        ws.Range("T3").Value = "モード項目なし"
        Exit Sub
    End If
    
    ' データを配列に読み込み
    Dim dataArr As Variant
    dataArr = ws.Range("AG13:AG" & lastRow).Value
    
    ' 除外項目の辞書
    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.Add "A", True
    excludeDict.Add "B", True
    excludeDict.Add "C", True
    excludeDict.Add "D", True
    excludeDict.Add "E", True
    excludeDict.Add "Fr RH", True
    
    ' 有効なモード項目を収集
    Dim modeItems As Object
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim cellValue As String
    
    For i = 1 To UBound(dataArr, 1)
        cellValue = Trim(CStr(dataArr(i, 1)))
        If cellValue <> "" And Not excludeDict.Exists(cellValue) And Not modeItems.Exists(cellValue) Then
            modeItems.Add cellValue, cellValue
        End If
    Next i
    
    ' 入力規則設定
    If modeItems.Count > 0 Then
        Dim modeList As String
        modeList = Join(modeItems.Keys, ",")
        
        With ws.Range("T3")
            .Validation.Delete
            .Value = ""
            .Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=modeList
        End With
    Else
        ws.Range("T3").Validation.Delete
        ws.Range("T3").Value = "モード項目なし"
    End If
End Sub