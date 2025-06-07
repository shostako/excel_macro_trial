Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Sub グラフ表示_ゾーンFR流出_最適化版()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 作成日: 2025/06/07
    ' 最適化版：処理速度向上と進捗表示の詳細化
    
    Dim ws As Worksheet
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String ' E3: 発生
    Dim discovery2Value As String ' E4: 発見2
    Dim arrDiscovery2 As Variant
    Dim i As Long
    Dim isProcessing As Boolean    ' 「発生」が「加工」工程判定用
    Dim isMould As Boolean         ' 「発生」が「モール」工程判定用
    Dim isDiscovery2Empty As Boolean ' 発見2の値が空か判定用
    Dim commentText As String      ' D6に設定するコメント用
    
    ' エラー処理を設定
    On Error GoTo ErrorHandler
    
    ' 高速化の三種の神器（最優先）
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
    
    ' 日付範囲の取得（セルE1～E2）
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
    
    ' 「発生」が「加工」または「モール」かどうかを判定
    isProcessing = (occurrenceValue = "加工")
    isMould = (occurrenceValue = "モール")
    
    ' 発見2値をカンマ区切りで配列に分割
    If Not isDiscovery2Empty Then
        arrDiscovery2 = Split(discovery2Value, ",")
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            arrDiscovery2(i) = Trim(arrDiscovery2(i))
        Next i
    End If
    
    ' 全ピボットテーブルの手動更新モードを設定（一括更新のため）
    Application.StatusBar = "ピボットテーブルの更新モードを設定中..."
    pt1.ManualUpdate = True
    pt2.ManualUpdate = True
    pt3.ManualUpdate = True
    pt4.ManualUpdate = True
    pt5.ManualUpdate = True
    
    ' モード2フィルタをリセット
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(Array(pt1, pt2, pt3, pt4, pt5))
    
    ' 各ピボットテーブルのフィルタ設定（個別進捗表示）
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
    
    ' ピボットテーブルの一括更新
    Application.StatusBar = "ピボットテーブルを更新中..."
    pt1.ManualUpdate = False
    pt2.ManualUpdate = False
    pt3.ManualUpdate = False
    pt4.ManualUpdate = False
    pt5.ManualUpdate = False
    
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
        ' その他の場合
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
    Dim maxVals() As Double
    ReDim maxVals(1 To 4)
    
    ' 各ピボットテーブルから最大値を配列で取得（高速化）
    maxVals(1) = GetPivotTableMaxValueFast(pt1)
    maxVals(2) = GetPivotTableMaxValueFast(pt2)
    maxVals(3) = GetPivotTableMaxValueFast(pt3)
    maxVals(4) = GetPivotTableMaxValueFast(pt4)
    
    ' 全体の最大値を決定
    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxVals)
    
    ' 良い感じの軸最大値を計算
    Dim axisMax As Double
    Dim tickInterval As Double
    axisMax = GetNiceMaxValueOptimized(overallMax)
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
    Application.StatusBar = False
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
                                     ByVal discovery2Arr As Variant, _
                                     ByVal isDiscovery2Empty As Boolean)
    ' 最適化されたピボットテーブルフィルタリング
    
    Dim pi As PivotItem
    Dim dateDict As Object
    Dim i As Long
    
    On Error Resume Next
    
    ' 日付フィールドの最適化（Dictionary使用）
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    With pt.PivotFields("日付")
        .ClearAllFilters
        ' 有効な日付を事前に判定
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                Dim d As Date
                d = CDate(pi.Name)
                dateDict(pi.Name) = (d >= startDate And d <= endDate)
            Else
                dateDict(pi.Name) = False
            End If
        Next pi
        
        ' 一括で表示/非表示を設定
        For Each pi In .PivotItems
            pi.Visible = dateDict(pi.Name)
        Next pi
    End With
    
    ' アル/ノア フィールド
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    
    ' Fr/Rr フィールド
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    
    ' 発生 フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2 フィールド（Dictionary使用で高速化）
    If Not isDiscovery2Empty Then
        Dim discovery2Dict As Object
        Set discovery2Dict = CreateObject("Scripting.Dictionary")
        
        ' 配列の値を辞書に格納
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

Private Sub FilterPivotTableForModeOptimized(ByVal pt As PivotTable, _
                                            ByVal startDate As Date, _
                                            ByVal endDate As Date, _
                                            ByVal occurrenceFilter As String, _
                                            ByVal discovery2Arr As Variant, _
                                            ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用ピボットテーブル専用の最適化フィルタリング
    
    Dim pi As PivotItem
    Dim dateDict As Object
    Dim i As Long
    
    On Error Resume Next
    
    ' 日付フィールドの最適化
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    With pt.PivotFields("日付")
        .ClearAllFilters
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                Dim d As Date
                d = CDate(pi.Name)
                dateDict(pi.Name) = (d >= startDate And d <= endDate)
            Else
                dateDict(pi.Name) = False
            End If
        Next pi
        
        For Each pi In .PivotItems
            pi.Visible = dateDict(pi.Name)
        Next pi
    End With
    
    ' アル/ノア・Fr/Rr：全て表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド
    If Not isDiscovery2Empty Then
        Dim discovery2Dict As Object
        Set discovery2Dict = CreateObject("Scripting.Dictionary")
        
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

Private Sub ResetMode2Filters(ByVal pivotTables As Variant)
    ' モード2フィルタを一括リセット
    Dim pt As PivotTable
    Dim i As Long
    
    On Error Resume Next
    For i = LBound(pivotTables) To UBound(pivotTables)
        Set pt = pivotTables(i)
        With pt.PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = "(すべて)"
        End With
    Next i
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByVal chartNames As Variant, ByVal visibilities As Variant)
    ' グラフの表示/非表示を一括設定
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

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を配列で高速取得
    Dim maxVal As Double
    Dim dataRange As Range
    Dim arr As Variant
    Dim i As Long, j As Long
    
    On Error Resume Next
    Set dataRange = pt.DataBodyRange
    
    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' 配列に一括読み込み
    arr = dataRange.Value
    
    maxVal = 0
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
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
    
    Dim targetValue As Double
    targetValue = maxValue * 1.15 ' 15%の余裕
    
    ' 切りの良い数字に調整
    Select Case True
        Case targetValue <= 10
            GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(targetValue, 1)
        Case targetValue <= 50
            GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(targetValue, 5)
        Case targetValue <= 100
            GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(targetValue, 10)
        Case targetValue <= 500
            GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(targetValue, 50)
        Case targetValue <= 1000
            GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(targetValue, 100)
        Case Else
            Dim magnitude As Long
            magnitude = 10 ^ Int(Log(targetValue) / Log(10))
            GetNiceMaxValueOptimized = Application.WorksheetFunction.Ceiling(targetValue, magnitude)
    End Select
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 軸の最大値に基づいて適切な目盛り間隔を計算
    
    Dim targetTicks As Long
    Dim roughInterval As Double
    
    targetTicks = 6
    roughInterval = maxValue / targetTicks
    
    ' 切りの良い間隔に調整
    Select Case True
        Case roughInterval <= 1
            GetNiceTickInterval = 1
        Case roughInterval <= 2
            GetNiceTickInterval = 2
        Case roughInterval <= 5
            GetNiceTickInterval = 5
        Case roughInterval <= 10
            GetNiceTickInterval = 10
        Case roughInterval <= 20
            GetNiceTickInterval = 20
        Case roughInterval <= 50
            GetNiceTickInterval = 50
        Case roughInterval <= 100
            GetNiceTickInterval = 100
        Case Else
            Dim magnitude As Long
            magnitude = 10 ^ Int(Log(roughInterval) / Log(10))
            GetNiceTickInterval = Application.WorksheetFunction.Ceiling(roughInterval / magnitude, 1) * magnitude
    End Select
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, ByVal chartName As String, ByVal maxValue As Double, ByVal tickInterval As Double)
    ' グラフの縦軸設定
    Dim chObj As ChartObject
    Dim ch As Chart
    
    On Error Resume Next
    Set chObj = ws.ChartObjects(chartName)
    
    If Not chObj Is Nothing Then
        Set ch = chObj.Chart
        
        With ch.Axes(xlValue)
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
    Set ch = Nothing
    On Error GoTo 0
End Sub

Private Sub SetModeFieldValidation(ByVal ws As Worksheet)
    ' モードフィールドの入力規則を高速設定
    
    Dim modeItems As Object
    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long
    Dim modeList As String
    Dim excludeDict As Object
    
    On Error Resume Next
    
    ' 除外値の辞書
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.Add "A", True
    excludeDict.Add "B", True
    excludeDict.Add "C", True
    excludeDict.Add "D", True
    excludeDict.Add "E", True
    excludeDict.Add "Fr RH", True
    
    ' モード項目の辞書
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    ' AG列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        ' 配列で一括読み込み（高速化）
        arr = ws.Range("AG13:AG" & lastRow).Value
        
        For i = LBound(arr, 1) To UBound(arr, 1)
            Dim cellValue As String
            cellValue = Trim(CStr(arr(i, 1)))
            
            If cellValue <> "" And Not excludeDict.Exists(cellValue) And Not modeItems.Exists(cellValue) Then
                modeItems.Add cellValue, cellValue
            End If
        Next i
    End If
    
    ' 入力規則設定
    If modeItems.Count > 0 Then
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
    
    On Error GoTo 0
End Sub