Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Option Explicit

Sub グラフ表示_ゾーンFR流出_最適化版()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 最適化版: 処理速度向上とステータスバー詳細表示
    ' 作成日: 2025/06/07
    
    Dim ws As Worksheet
    Dim pivotTables(1 To 5) As PivotTable
    Dim pivotNames As Variant
    Dim pivotDescriptions As Variant
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String
    Dim discovery2Value As String
    Dim arrDiscovery2 As Variant
    Dim isProcessing As Boolean
    Dim isMould As Boolean
    Dim isDiscovery2Empty As Boolean
    Dim commentText As String
    Dim i As Long
    
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
    
    ' ピボットテーブル名と説明の配列
    pivotNames = Array("ピボットテーブル31", "ピボットテーブル32", "ピボットテーブル33", _
                      "ピボットテーブル34", "ピボットテーブル35")
    pivotDescriptions = Array("アルヴェル Fr", "アルヴェル Rr", "ノアヴォク Fr", _
                            "ノアヴォク Rr", "モード抽出用")
    
    ' ピボットテーブルの取得と検証
    Application.StatusBar = "ピボットテーブルを確認中..."
    For i = 1 To 5
        On Error Resume Next
        Set pivotTables(i) = ws.PivotTables(pivotNames(i - 1))
        On Error GoTo ErrorHandler
        
        If pivotTables(i) Is Nothing Then
            MsgBox "ピボットテーブル '" & pivotNames(i - 1) & "' が見つかりません。", vbExclamation
            GoTo Cleanup
        End If
    Next i
    
    ' パラメータの取得
    Application.StatusBar = "パラメータを読み込み中..."
    If Not GetParameters(ws, dtStart, dtEnd, occurrenceValue, discovery2Value, _
                        isProcessing, isMould, isDiscovery2Empty, arrDiscovery2) Then
        GoTo Cleanup
    End If
    
    ' 全ピボットテーブルの更新を一時停止
    Application.StatusBar = "ピボットテーブルの準備中..."
    For i = 1 To 5
        pivotTables(i).ManualUpdate = True
    Next i
    
    ' モード2フィルタを一括リセット
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(pivotTables)
    
    ' ピボットテーブルのフィルタ設定（個別に進捗表示）
    For i = 1 To 4
        Application.StatusBar = pivotDescriptions(i - 1) & " ピボットテーブルを設定中..."
        
        Select Case i
            Case 1
                Call FilterPivotTableOptimized(pivotTables(1), dtStart, dtEnd, "アルヴェル", "Fr", _
                                              occurrenceValue, arrDiscovery2, isDiscovery2Empty)
            Case 2
                Call FilterPivotTableOptimized(pivotTables(2), dtStart, dtEnd, "アルヴェル", "Rr", _
                                              occurrenceValue, arrDiscovery2, isDiscovery2Empty)
            Case 3
                Call FilterPivotTableOptimized(pivotTables(3), dtStart, dtEnd, "ノアヴォク", "Fr", _
                                              occurrenceValue, arrDiscovery2, isDiscovery2Empty)
            Case 4
                Call FilterPivotTableOptimized(pivotTables(4), dtStart, dtEnd, "ノアヴォク", "Rr", _
                                              occurrenceValue, arrDiscovery2, isDiscovery2Empty)
        End Select
    Next i
    
    ' モード抽出用ピボットテーブルの設定
    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call FilterPivotTableForModeOptimized(pivotTables(5), dtStart, dtEnd, _
                                         occurrenceValue, arrDiscovery2, isDiscovery2Empty)
    
    ' 全ピボットテーブルを一括更新
    Application.StatusBar = "ピボットテーブルを更新中..."
    For i = 1 To 5
        pivotTables(i).ManualUpdate = False
        pivotTables(i).RefreshTable
    Next i
    
    ' グラフ表示設定
    Application.StatusBar = "グラフ表示を設定中..."
    Call SetGraphVisibility(ws, isProcessing, isMould)
    
    ' グラフ軸の動的調整
    Application.StatusBar = "グラフ軸を調整中..."
    Call AdjustChartAxes(ws, pivotTables, isProcessing, isMould)
    
    ' コメント設定
    Application.StatusBar = "コメントを設定中..."
    commentText = GetCommentText(isProcessing, isMould, occurrenceValue, dtStart, dtEnd)
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの入力規則設定
    Application.StatusBar = "モードフィールドを設定中..."
    Call SetModeFieldValidation(ws)

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
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "マクロエラー"
    Resume Cleanup
End Sub

Private Function GetParameters(ByVal ws As Worksheet, _
                              ByRef dtStart As Date, _
                              ByRef dtEnd As Date, _
                              ByRef occurrenceValue As String, _
                              ByRef discovery2Value As String, _
                              ByRef isProcessing As Boolean, _
                              ByRef isMould As Boolean, _
                              ByRef isDiscovery2Empty As Boolean, _
                              ByRef arrDiscovery2 As Variant) As Boolean
    ' パラメータを取得し、検証する
    
    GetParameters = False
    
    ' 日付範囲の取得
    If IsDate(ws.Range("E1").Value) And IsDate(ws.Range("E2").Value) Then
        dtStart = ws.Range("E1").Value
        dtEnd = ws.Range("E2").Value
    Else
        MsgBox "日付範囲が正しく設定されていません。セルE1とE2を確認してください。", vbExclamation
        Exit Function
    End If
    
    ' 発生値と発見2値の取得
    occurrenceValue = Trim(CStr(ws.Range("E3").Value))
    discovery2Value = Trim(CStr(ws.Range("E4").Value))
    
    ' 発生値のエラーチェック
    If occurrenceValue = "" Then
        MsgBox "発生の値が設定されていません。セルE3を確認してください。", vbExclamation
        Exit Function
    End If
    
    ' フラグの設定
    isDiscovery2Empty = (discovery2Value = "")
    isProcessing = (occurrenceValue = "加工")
    isMould = (occurrenceValue = "モール")
    
    ' 発見2値をカンマ区切りで配列に分割
    If Not isDiscovery2Empty Then
        arrDiscovery2 = Split(discovery2Value, ",")
        Dim i As Long
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            arrDiscovery2(i) = Trim(arrDiscovery2(i))
        Next i
    End If
    
    GetParameters = True
End Function

Private Sub ResetMode2Filters(ByRef pivotTables() As PivotTable)
    ' モード2フィルタを一括リセット
    Dim i As Long
    Dim pf As PivotField
    
    For i = 1 To 5
        On Error Resume Next
        Set pf = pivotTables(i).PivotFields("モード2")
        If Not pf Is Nothing Then
            pf.ClearAllFilters
            pf.CurrentPage = "(すべて)"
        End If
        On Error GoTo 0
    Next i
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
    
    ' 日付フィールドのフィルタリング（配列処理で高速化）
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim pi As PivotItem
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                pi.Visible = (CDate(pi.Name) >= startDate And CDate(pi.Name) <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' ページフィールドの設定
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールドのフィルタリング（Dictionary使用で高速化）
    If Not isDiscovery2Empty Then
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' 配列の値を辞書に格納
        Dim i As Long
        For i = LBound(discovery2Arr) To UBound(discovery2Arr)
            dict(discovery2Arr(i)) = True
        Next i
        
        ' フィルタ適用
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = dict.Exists(pi.Name)
            Next pi
        End With
        
        Set dict = Nothing
    End If
    
    On Error GoTo 0
End Sub

Private Sub FilterPivotTableForModeOptimized(ByVal pt As PivotTable, _
                                            ByVal startDate As Date, _
                                            ByVal endDate As Date, _
                                            ByVal occurrenceFilter As String, _
                                            ByVal discovery2Arr As Variant, _
                                            ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用の最適化されたフィルタリング
    
    On Error Resume Next
    
    ' 日付フィールドのフィルタリング
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim pi As PivotItem
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                pi.Visible = (CDate(pi.Name) >= startDate And CDate(pi.Name) <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' アル/ノア、Fr/Rrは全て表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールドのフィルタリング
    If Not isDiscovery2Empty Then
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        For i = LBound(discovery2Arr) To UBound(discovery2Arr)
            dict(discovery2Arr(i)) = True
        Next i
        
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = dict.Exists(pi.Name)
            Next pi
        End With
        
        Set dict = Nothing
    End If
    
    On Error GoTo 0
End Sub

Private Sub SetGraphVisibility(ByVal ws As Worksheet, _
                              ByVal isProcessing As Boolean, _
                              ByVal isMould As Boolean)
    ' グラフの表示/非表示を一括設定
    
    Dim showGraph(1 To 4) As Boolean
    
    If isProcessing Then
        ' 加工の場合、全て非表示
        showGraph(1) = False
        showGraph(2) = False
        showGraph(3) = False
        showGraph(4) = False
    ElseIf isMould Then
        ' モールの場合、1と2のみ表示
        showGraph(1) = True
        showGraph(2) = True
        showGraph(3) = False
        showGraph(4) = False
    Else
        ' その他の場合、全て表示
        showGraph(1) = True
        showGraph(2) = True
        showGraph(3) = True
        showGraph(4) = True
    End If
    
    ' 一括設定
    Call SetChartVisibilityBatch(ws, showGraph)
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByRef showGraph() As Boolean)
    ' グラフの表示/非表示を一括で設定
    Dim i As Long
    Dim chObj As ChartObject
    
    On Error Resume Next
    For i = 1 To 4
        Set chObj = ws.ChartObjects("グラフ" & i)
        If Not chObj Is Nothing Then
            chObj.Visible = showGraph(i)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub AdjustChartAxes(ByVal ws As Worksheet, _
                           ByRef pivotTables() As PivotTable, _
                           ByVal isProcessing As Boolean, _
                           ByVal isMould As Boolean)
    ' グラフ軸の動的調整
    
    If isProcessing Then Exit Sub ' 加工の場合はグラフ非表示なので調整不要
    
    ' 各ピボットテーブルから最大値を取得（配列使用で高速化）
    Dim maxValues(1 To 4) As Double
    Dim i As Long
    
    For i = 1 To 4
        maxValues(i) = GetPivotTableMaxValueFast(pivotTables(i))
    Next i
    
    ' 全体の最大値を決定
    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxValues)
    
    ' 良い感じの軸最大値と目盛り間隔を計算
    Dim axisMax As Double
    Dim tickInterval As Double
    axisMax = GetNiceMaxValue(overallMax)
    tickInterval = GetNiceTickInterval(axisMax)
    
    ' 各グラフに適用
    On Error Resume Next
    For i = 1 To 4
        If (Not isMould) Or (isMould And i <= 2) Then
            Call SetChartAxisSettings(ws, "グラフ" & i, axisMax, tickInterval)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を高速に取得（配列使用）
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

Private Function GetNiceMaxValue(ByVal maxValue As Double) As Double
    ' 適切な軸の最大値を計算
    If maxValue <= 0 Then
        GetNiceMaxValue = 10
        Exit Function
    End If
    
    ' 最大値の110%～120%で最も近い切りの良い数字
    Dim targetValue As Double
    targetValue = maxValue * 1.15
    
    ' 桁数に応じて丸める
    Select Case True
        Case targetValue <= 10
            GetNiceMaxValue = Application.WorksheetFunction.Ceiling(targetValue, 1)
        Case targetValue <= 50
            GetNiceMaxValue = Application.WorksheetFunction.Ceiling(targetValue, 5)
        Case targetValue <= 100
            GetNiceMaxValue = Application.WorksheetFunction.Ceiling(targetValue, 10)
        Case targetValue <= 500
            GetNiceMaxValue = Application.WorksheetFunction.Ceiling(targetValue, 50)
        Case targetValue <= 1000
            GetNiceMaxValue = Application.WorksheetFunction.Ceiling(targetValue, 100)
        Case Else
            GetNiceMaxValue = Application.WorksheetFunction.Ceiling(targetValue, 500)
    End Select
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 適切な目盛り間隔を計算
    Dim targetTicks As Long
    targetTicks = 6
    
    Dim roughInterval As Double
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
            GetNiceTickInterval = Application.WorksheetFunction.Ceiling(roughInterval, 50)
    End Select
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, _
                                ByVal chartName As String, _
                                ByVal maxValue As Double, _
                                ByVal tickInterval As Double)
    ' グラフの軸設定
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

Private Function GetCommentText(ByVal isProcessing As Boolean, _
                               ByVal isMould As Boolean, _
                               ByVal occurrenceValue As String, _
                               ByVal dtStart As Date, _
                               ByVal dtEnd As Date) As String
    ' コメントテキストを生成
    If isProcessing Then
        GetCommentText = "発生が「加工」のため、グラフは表示されません。"
    Else
        Dim startDateStr As String
        Dim endDateStr As String
        startDateStr = Format(dtStart, "m/d")
        endDateStr = Format(dtEnd, "m/d")
        GetCommentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    End If
End Function

Private Sub SetModeFieldValidation(ByVal ws As Worksheet)
    ' モードフィールドの入力規則を設定
    Dim modeItems As Object
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim cellValue As String
    Dim excludeList As Variant
    Dim i As Long
    
    ' 除外リスト
    excludeList = Array("A", "B", "C", "D", "E", "Fr RH")
    
    ' Dictionary使って重複排除
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    ' AG列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        ' AG13以降のセルを配列で高速処理
        Dim arr As Variant
        arr = ws.Range("AG13:AG" & lastRow).Value
        
        For i = 1 To UBound(arr, 1)
            cellValue = Trim(CStr(arr(i, 1)))
            
            If cellValue <> "" And Not IsInArray(cellValue, excludeList) Then
                modeItems(cellValue) = True
            End If
        Next i
    End If
    
    ' リスト文字列作成と入力規則設定
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
    
    Set modeItems = Nothing
End Sub

Private Function IsInArray(ByVal searchValue As String, ByRef arr As Variant) As Boolean
    ' 配列内に値が存在するかチェック
    Dim element As Variant
    For Each element In arr
        If element = searchValue Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function