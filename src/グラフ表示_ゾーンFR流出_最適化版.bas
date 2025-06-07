Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Sub グラフ表示_ゾーンFR流出_最適化版()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 作成日: 2025/06/07
    ' 最適化版: 処理速度向上とステータスバー詳細表示
    
    Dim ws As Worksheet
    Dim pivotTables(1 To 5) As PivotTable
    Dim pivotNames As Variant
    Dim pivotDescriptions As Variant
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String
    Dim discovery2Value As String
    Dim discovery2Dict As Object
    Dim i As Long
    Dim isProcessing As Boolean
    Dim isMould As Boolean
    Dim isDiscovery2Empty As Boolean
    Dim commentText As String
    
    ' エラー処理を設定
    On Error GoTo ErrorHandler
    
    ' 高速化の三種の神器
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "初期化中..."
    
    ' ワークシートの取得
    Set ws = ThisWorkbook.Worksheets("ゾーンFrRr流出")
    
    ' ピボットテーブル名と説明を配列で管理
    pivotNames = Array("ピボットテーブル31", "ピボットテーブル32", "ピボットテーブル33", _
                      "ピボットテーブル34", "ピボットテーブル35")
    pivotDescriptions = Array("アルヴェル Fr", "アルヴェル Rr", "ノアヴォク Fr", _
                             "ノアヴォク Rr", "モード抽出用")
    
    ' ピボットテーブルの取得と検証
    For i = 1 To 5
        On Error Resume Next
        Set pivotTables(i) = ws.PivotTables(pivotNames(i - 1))
        If Err.Number <> 0 Then
            MsgBox "ピボットテーブル '" & pivotNames(i - 1) & "' が見つかりません。", vbExclamation
            GoTo Cleanup
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' 入力値の取得と検証
    If Not IsDate(ws.Range("E1").Value) Or Not IsDate(ws.Range("E2").Value) Then
        MsgBox "日付範囲が正しく設定されていません。セルE1とE2を確認してください。", vbExclamation
        GoTo Cleanup
    End If
    
    dtStart = ws.Range("E1").Value
    dtEnd = ws.Range("E2").Value
    occurrenceValue = Trim(CStr(ws.Range("E3").Value))
    discovery2Value = Trim(CStr(ws.Range("E4").Value))
    
    If occurrenceValue = "" Then
        MsgBox "発生の値が設定されていません。セルE3を確認してください。", vbExclamation
        GoTo Cleanup
    End If
    
    ' 条件判定
    isDiscovery2Empty = (discovery2Value = "")
    isProcessing = (occurrenceValue = "加工")
    isMould = (occurrenceValue = "モール")
    
    ' 発見2値をDictionaryで管理（高速化）
    Set discovery2Dict = CreateObject("Scripting.Dictionary")
    If Not isDiscovery2Empty Then
        Dim arrDiscovery2 As Variant
        arrDiscovery2 = Split(discovery2Value, ",")
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            discovery2Dict(Trim(arrDiscovery2(i))) = True
        Next i
    End If
    
    ' 全ピボットテーブルの手動更新モードを設定（一括処理で高速化）
    Application.StatusBar = "ピボットテーブルを準備中..."
    For i = 1 To 5
        pivotTables(i).ManualUpdate = True
    Next i
    
    ' モード2フィルタをリセット
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(pivotTables)
    
    ' 各ピボットテーブルの設定（個別進捗表示）
    Application.StatusBar = pivotDescriptions(0) & " ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pivotTables(1), dtStart, dtEnd, "アルヴェル", "Fr", _
                                  occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = pivotDescriptions(1) & " ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pivotTables(2), dtStart, dtEnd, "アルヴェル", "Rr", _
                                  occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = pivotDescriptions(2) & " ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pivotTables(3), dtStart, dtEnd, "ノアヴォク", "Fr", _
                                  occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = pivotDescriptions(3) & " ピボットテーブルを設定中..."
    Call FilterPivotTableOptimized(pivotTables(4), dtStart, dtEnd, "ノアヴォク", "Rr", _
                                  occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = pivotDescriptions(4) & " ピボットテーブルを設定中..."
    Call FilterPivotTableForModeOptimized(pivotTables(5), dtStart, dtEnd, occurrenceValue, _
                                         discovery2Dict, isDiscovery2Empty)
    
    ' 一括でピボットテーブルを更新（最も効率的）
    Application.StatusBar = "ピボットテーブルを更新中..."
    For i = 1 To 5
        pivotTables(i).ManualUpdate = False
        pivotTables(i).RefreshTable
    Next i
    
    ' グラフ表示設定
    Application.StatusBar = "グラフ表示を設定中..."
    Dim chartVisibility(1 To 4) As Boolean
    
    Select Case True
        Case isProcessing
            ' 加工の場合：全グラフ非表示
            chartVisibility(1) = False
            chartVisibility(2) = False
            chartVisibility(3) = False
            chartVisibility(4) = False
            commentText = "発生が「加工」のため、グラフは表示されません。"
            
        Case isMould
            ' モールの場合：グラフ1,2のみ表示
            chartVisibility(1) = True
            chartVisibility(2) = True
            chartVisibility(3) = False
            chartVisibility(4) = False
            commentText = occurrenceValue & " 流出不良集計 " & Format(dtStart, "m/d") & _
                         " ～ " & Format(dtEnd, "m/d")
            
        Case Else
            ' その他の場合：全グラフ表示
            chartVisibility(1) = True
            chartVisibility(2) = True
            chartVisibility(3) = True
            chartVisibility(4) = True
            commentText = occurrenceValue & " 流出不良集計 " & Format(dtStart, "m/d") & _
                         " ～ " & Format(dtEnd, "m/d")
    End Select
    
    ' グラフ表示/非表示を一括設定
    Call SetChartVisibilityBatch(ws, chartVisibility)
    
    ' グラフ軸の動的調整
    Application.StatusBar = "グラフ軸を調整中..."
    Dim maxValues(1 To 4) As Double
    Dim overallMax As Double
    Dim axisMax As Double
    Dim tickInterval As Double
    
    ' 各ピボットテーブルから最大値を高速取得
    For i = 1 To 4
        maxValues(i) = GetPivotTableMaxValueFast(pivotTables(i))
    Next i
    
    ' 全体の最大値を決定
    overallMax = Application.WorksheetFunction.Max(maxValues)
    
    ' 適切な軸設定を計算
    axisMax = GetNiceMaxValue(overallMax)
    tickInterval = GetNiceTickInterval(axisMax)
    
    ' 表示されているグラフのみ軸設定を適用
    For i = 1 To 4
        If chartVisibility(i) Then
            Call SetChartAxisSettings(ws, "グラフ" & i, axisMax, tickInterval)
        End If
    Next i
    
    ' D6にコメントを設定
    Application.StatusBar = "最終処理中..."
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの入力規則設定（高速化版）
    Call SetModeValidation(ws)
    
Cleanup:
    Application.StatusBar = "処理が完了しました。"
    Application.Wait Now + TimeValue("00:00:01")
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' オブジェクトの解放
    Set ws = Nothing
    For i = 1 To 5
        Set pivotTables(i) = Nothing
    Next i
    Set discovery2Dict = Nothing
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
    
    Dim pi As PivotItem
    Dim d As Date
    
    On Error Resume Next
    
    ' 日付フィールド（高速化：条件を直接評価）
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
    
    ' レポートフィルタ（CurrentPageで高速設定）
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド（Dictionary使用で高速化）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub FilterPivotTableForModeOptimized(ByVal pt As PivotTable, _
                                           ByVal startDate As Date, _
                                           ByVal endDate As Date, _
                                           ByVal occurrenceFilter As String, _
                                           ByVal discovery2Dict As Object, _
                                           ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用ピボットテーブル専用フィルタリング（最適化版）
    
    Dim pi As PivotItem
    Dim d As Date
    
    On Error Resume Next
    
    ' 日付フィールド
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
    
    ' 発見2フィールド
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub ResetMode2Filters(ByRef pivotTables() As PivotTable)
    ' モード2フィルタの一括リセット
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To 5
        With pivotTables(i).PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = "(すべて)"
        End With
    Next i
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByRef visibility() As Boolean)
    ' グラフ表示/非表示の一括設定
    Dim i As Long
    Dim chObj As ChartObject
    
    On Error Resume Next
    For i = 1 To 4
        Set chObj = ws.ChartObjects("グラフ" & i)
        If Not chObj Is Nothing Then
            chObj.Visible = visibility(i)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を高速取得（配列使用）
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim maxVal As Double
    Dim i As Long, j As Long
    
    On Error Resume Next
    Set dataRange = pt.DataBodyRange
    
    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' 配列に一括読み込み（高速化）
    dataArray = dataRange.Value
    
    maxVal = 0
    For i = 1 To UBound(dataArray, 1)
        For j = 1 To UBound(dataArray, 2)
            If IsNumeric(dataArray(i, j)) And dataArray(i, j) > maxVal Then
                maxVal = dataArray(i, j)
            End If
        Next j
    Next i
    
    GetPivotTableMaxValueFast = maxVal
    On Error GoTo 0
End Function

Private Function GetNiceMaxValue(ByVal maxValue As Double) As Double
    ' データの最大値から適切な軸の最大値を計算
    Dim targetValue As Double
    Dim magnitude As Long
    Dim base As Double
    
    If maxValue <= 0 Then
        GetNiceMaxValue = 10
        Exit Function
    End If
    
    ' 最大値の115%を目標値とする
    targetValue = maxValue * 1.15
    
    ' 桁数を取得
    magnitude = Int(Log(targetValue) / Log(10))
    base = 10 ^ magnitude
    
    ' 切りの良い数値に調整
    Select Case targetValue / base
        Case Is <= 1.5
            GetNiceMaxValue = 1.5 * base
        Case Is <= 2
            GetNiceMaxValue = 2 * base
        Case Is <= 3
            GetNiceMaxValue = 3 * base
        Case Is <= 5
            GetNiceMaxValue = 5 * base
        Case Is <= 7
            GetNiceMaxValue = 7 * base
        Case Else
            GetNiceMaxValue = 10 * base
    End Select
    
    ' 最小でも最大値+1は保証
    If GetNiceMaxValue <= maxValue Then
        GetNiceMaxValue = maxValue + 1
    End If
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 軸の最大値に基づいて適切な目盛り間隔を計算
    Dim targetTicks As Long
    Dim roughInterval As Double
    
    targetTicks = 6
    roughInterval = maxValue / targetTicks
    
    ' 切りの良い間隔に調整（簡潔な実装）
    Select Case True
        Case roughInterval <= 1: GetNiceTickInterval = 1
        Case roughInterval <= 2: GetNiceTickInterval = 2
        Case roughInterval <= 5: GetNiceTickInterval = 5
        Case roughInterval <= 10: GetNiceTickInterval = 10
        Case roughInterval <= 20: GetNiceTickInterval = 20
        Case roughInterval <= 25: GetNiceTickInterval = 25
        Case roughInterval <= 50: GetNiceTickInterval = 50
        Case roughInterval <= 100: GetNiceTickInterval = 100
        Case Else
            Dim magnitude As Long
            magnitude = Int(Log(roughInterval) / Log(10))
            Dim base As Double
            base = 10 ^ magnitude
            
            Select Case roughInterval / base
                Case Is <= 2: GetNiceTickInterval = 2 * base
                Case Is <= 5: GetNiceTickInterval = 5 * base
                Case Else: GetNiceTickInterval = 10 * base
            End Select
    End Select
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, ByVal chartName As String, _
                               ByVal maxValue As Double, ByVal tickInterval As Double)
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
    
    On Error GoTo 0
End Sub

Private Sub SetModeValidation(ByVal ws As Worksheet)
    ' モードフィールドの入力規則設定（高速化版）
    Dim modeItems As Object
    Dim rng As Range
    Dim dataArray As Variant
    Dim excludeList As Variant
    Dim i As Long
    Dim cellValue As String
    
    Set modeItems = CreateObject("Scripting.Dictionary")
    excludeList = Array("A", "B", "C", "D", "E", "Fr RH")
    
    ' 除外リストをDictionaryに変換（高速検索用）
    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(excludeList)
        excludeDict(excludeList(i)) = True
    Next i
    
    ' AG列のデータを配列で一括取得（高速化）
    On Error Resume Next
    Set rng = ws.Range("AG13:AG" & ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row)
    
    If rng.Row >= 13 Then
        dataArray = rng.Value
        
        ' 配列をループ（セル単位より高速）
        For i = 1 To UBound(dataArray, 1)
            cellValue = Trim(CStr(dataArray(i, 1)))
            
            If cellValue <> "" And Not excludeDict.Exists(cellValue) And _
               Not modeItems.Exists(cellValue) Then
                modeItems.Add cellValue, True
            End If
        Next i
    End If
    
    ' 入力規則を設定
    If modeItems.Count > 0 Then
        With ws.Range("T3")
            .Validation.Delete
            .Value = ""
            .Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=Join(modeItems.Keys, ",")
        End With
    Else
        ws.Range("T3").Validation.Delete
        ws.Range("T3").Value = "モード項目なし"
    End If
    
    On Error GoTo 0
End Sub