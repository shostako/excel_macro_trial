Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Option Explicit

Sub グラフ表示_ゾーンFR流出_最適化版()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 作成日: 2025/06/07
    ' 最適化版: 処理速度向上と詳細な進捗表示を実装

    Dim ws As Worksheet
    Dim pivotTables(1 To 5) As PivotTable
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String ' E3: 発生
    Dim discovery2Value As String ' E4: 発見2
    Dim discovery2Dict As Object ' 発見2の高速検索用
    Dim isProcessing As Boolean
    Dim isMould As Boolean
    Dim commentText As String
    
    ' エラー処理を設定
    On Error GoTo ErrorHandler

    ' 高速化設定（三種の神器）
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

    ' ピボットテーブルの一括取得
    Application.StatusBar = "ピボットテーブルを読み込み中..."
    Set pivotTables(1) = ws.PivotTables("ピボットテーブル31") ' アルヴェル Fr
    Set pivotTables(2) = ws.PivotTables("ピボットテーブル32") ' アルヴェル Rr
    Set pivotTables(3) = ws.PivotTables("ピボットテーブル33") ' ノアヴォク Fr
    Set pivotTables(4) = ws.PivotTables("ピボットテーブル34") ' ノアヴォク Rr
    Set pivotTables(5) = ws.PivotTables("ピボットテーブル35") ' モード抽出用

    ' 入力値の取得と検証
    If Not ValidateInputs(ws, dtStart, dtEnd, occurrenceValue, discovery2Value) Then
        GoTo Cleanup
    End If

    ' 発見2の辞書を作成（高速検索用）
    Set discovery2Dict = CreateDiscovery2Dictionary(discovery2Value)
    
    ' 処理タイプの判定
    isProcessing = (occurrenceValue = "加工")
    isMould = (occurrenceValue = "モール")

    ' 全ピボットテーブルの手動更新モードを設定
    Application.StatusBar = "ピボットテーブルの更新モードを設定中..."
    Dim i As Long
    For i = 1 To 5
        pivotTables(i).ManualUpdate = True
    Next i

    ' モード2フィルタの一括リセット
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(pivotTables)

    ' 各ピボットテーブルのフィルタ設定（更新なし）
    Dim tableNames As Variant
    tableNames = Array("アルヴェル Fr", "アルヴェル Rr", "ノアヴォク Fr", "ノアヴォク Rr", "モード抽出")
    
    For i = 1 To 5
        Application.StatusBar = tableNames(i - 1) & " ピボットテーブルを設定中..."
        
        If i <= 4 Then
            ' 通常のピボットテーブル（1-4）
            Call SetPivotFilters(pivotTables(i), dtStart, dtEnd, _
                                Split(tableNames(i - 1))(0), _
                                Split(tableNames(i - 1))(1), _
                                occurrenceValue, discovery2Dict)
        Else
            ' モード抽出用ピボットテーブル（5）
            Call SetPivotFiltersForMode(pivotTables(5), dtStart, dtEnd, _
                                       occurrenceValue, discovery2Dict)
        End If
    Next i

    ' 全ピボットテーブルを一括更新
    Application.StatusBar = "ピボットテーブルを一括更新中..."
    For i = 1 To 5
        pivotTables(i).ManualUpdate = False
        pivotTables(i).RefreshTable
    Next i

    ' グラフ表示設定（一括処理）
    Application.StatusBar = "グラフ表示を設定中..."
    Dim graphVisibility(1 To 4) As Boolean
    Call DetermineGraphVisibility(isProcessing, isMould, graphVisibility)
    Call SetChartVisibilityBatch(ws, graphVisibility)

    ' グラフ軸の動的調整
    Application.StatusBar = "グラフ軸を調整中..."
    Call AdjustChartAxes(ws, pivotTables, graphVisibility)
    
    ' コメント設定
    Application.StatusBar = "コメントを設定中..."
    commentText = GenerateCommentText(occurrenceValue, dtStart, dtEnd, isProcessing)
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの入力規則設定
    Application.StatusBar = "モードフィールドを設定中..."
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

Private Function ValidateInputs(ByVal ws As Worksheet, _
                               ByRef dtStart As Date, _
                               ByRef dtEnd As Date, _
                               ByRef occurrenceValue As String, _
                               ByRef discovery2Value As String) As Boolean
    ' 入力値の検証
    
    ValidateInputs = False
    
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
    
    ValidateInputs = True
End Function

Private Function CreateDiscovery2Dictionary(ByVal discovery2Value As String) As Object
    ' 発見2の値を辞書化（高速検索用）
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If discovery2Value <> "" Then
        Dim arr As Variant
        arr = Split(discovery2Value, ",")
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            dict(Trim(arr(i))) = True
        Next i
    End If
    
    Set CreateDiscovery2Dictionary = dict
End Function

Private Sub ResetMode2Filters(ByRef pivotTables() As PivotTable)
    ' モード2フィルタの一括リセット
    
    Dim i As Long
    For i = 1 To 5
        On Error Resume Next
        With pivotTables(i).PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = "(すべて)"
        End With
        On Error GoTo 0
    Next i
End Sub

Private Sub SetPivotFilters(ByVal pt As PivotTable, _
                           ByVal startDate As Date, _
                           ByVal endDate As Date, _
                           ByVal alNoahFilter As String, _
                           ByVal frRrFilter As String, _
                           ByVal occurrenceFilter As String, _
                           ByVal discovery2Dict As Object)
    ' ピボットテーブルのフィルタ設定（更新なし）
    
    Dim pi As PivotItem
    
    On Error Resume Next
    
    ' 日付フィールド（配列を使用した高速処理）
    With pt.PivotFields("日付")
        .ClearAllFilters
        Dim dateArray() As Boolean
        ReDim dateArray(1 To .PivotItems.Count)
        Dim idx As Long: idx = 0
        
        For Each pi In .PivotItems
            idx = idx + 1
            If IsDate(pi.Name) Then
                Dim d As Date: d = CDate(pi.Name)
                dateArray(idx) = (d >= startDate And d <= endDate)
            Else
                dateArray(idx) = False
            End If
        Next pi
        
        ' 一括適用
        idx = 0
        For Each pi In .PivotItems
            idx = idx + 1
            pi.Visible = dateArray(idx)
        Next pi
    End With
    
    ' その他のフィールド
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド（辞書を使用した高速処理）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If discovery2Dict.Count > 0 Then
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub SetPivotFiltersForMode(ByVal pt As PivotTable, _
                                  ByVal startDate As Date, _
                                  ByVal endDate As Date, _
                                  ByVal occurrenceFilter As String, _
                                  ByVal discovery2Dict As Object)
    ' モード抽出用ピボットテーブルのフィルタ設定
    
    Dim pi As PivotItem
    
    On Error Resume Next
    
    ' 日付フィールド
    With pt.PivotFields("日付")
        .ClearAllFilters
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                Dim d As Date: d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With
    
    ' アル/ノアとFr/Rrは全て表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If discovery2Dict.Count > 0 Then
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub DetermineGraphVisibility(ByVal isProcessing As Boolean, _
                                   ByVal isMould As Boolean, _
                                   ByRef visibility() As Boolean)
    ' グラフ表示設定の決定
    
    If isProcessing Then
        ' 加工の場合：全て非表示
        visibility(1) = False
        visibility(2) = False
        visibility(3) = False
        visibility(4) = False
    ElseIf isMould Then
        ' モールの場合：1,2のみ表示
        visibility(1) = True
        visibility(2) = True
        visibility(3) = False
        visibility(4) = False
    Else
        ' その他：全て表示
        visibility(1) = True
        visibility(2) = True
        visibility(3) = True
        visibility(4) = True
    End If
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByRef visibility() As Boolean)
    ' グラフ表示/非表示の一括設定
    
    Dim i As Long
    For i = 1 To 4
        On Error Resume Next
        ws.ChartObjects("グラフ" & i).Visible = visibility(i)
        On Error GoTo 0
    Next i
End Sub

Private Sub AdjustChartAxes(ByVal ws As Worksheet, _
                           ByRef pivotTables() As PivotTable, _
                           ByRef visibility() As Boolean)
    ' グラフ軸の動的調整
    
    ' 各ピボットテーブルの最大値を配列で取得
    Dim maxValues(1 To 4) As Double
    Dim i As Long
    
    For i = 1 To 4
        maxValues(i) = GetPivotTableMaxValueFast(pivotTables(i))
    Next i
    
    ' 全体の最大値
    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxValues)
    
    ' 軸設定の計算
    Dim axisMax As Double, tickInterval As Double
    axisMax = CalculateNiceMaxValue(overallMax)
    tickInterval = CalculateTickInterval(axisMax)
    
    ' 表示されているグラフのみ軸設定
    For i = 1 To 4
        If visibility(i) Then
            On Error Resume Next
            With ws.ChartObjects("グラフ" & i).Chart.Axes(xlValue)
                .MaximumScaleIsAuto = False
                .MaximumScale = axisMax
                .MinimumScaleIsAuto = False
                .MinimumScale = 0
                .MajorUnitIsAuto = False
                .MajorUnit = tickInterval
                .MinorUnitIsAuto = False
                .MinorUnit = tickInterval / 2
            End With
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を高速取得（配列使用）
    
    On Error Resume Next
    
    Dim dataRange As Range
    Set dataRange = pt.DataBodyRange
    
    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' 配列で一括取得
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    Dim maxVal As Double: maxVal = 0
    Dim i As Long, j As Long
    
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

Private Function CalculateNiceMaxValue(ByVal maxValue As Double) As Double
    ' 良い感じの軸最大値を計算
    
    If maxValue <= 0 Then
        CalculateNiceMaxValue = 10
        Exit Function
    End If
    
    Dim target As Double
    target = maxValue * 1.15 ' 15%の余裕
    
    ' 桁数に基づいて切り上げ
    Dim magnitude As Long
    magnitude = Int(Log(target) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude
    
    Select Case target / base
        Case Is <= 1: CalculateNiceMaxValue = base
        Case Is <= 2: CalculateNiceMaxValue = 2 * base
        Case Is <= 3: CalculateNiceMaxValue = 3 * base
        Case Is <= 4: CalculateNiceMaxValue = 4 * base
        Case Is <= 5: CalculateNiceMaxValue = 5 * base
        Case Is <= 6: CalculateNiceMaxValue = 6 * base
        Case Is <= 7: CalculateNiceMaxValue = 7 * base
        Case Is <= 8: CalculateNiceMaxValue = 8 * base
        Case Is <= 9: CalculateNiceMaxValue = 9 * base
        Case Else: CalculateNiceMaxValue = 10 * base
    End Select
End Function

Private Function CalculateTickInterval(ByVal maxValue As Double) As Double
    ' 適切な目盛り間隔を計算
    
    Dim roughInterval As Double
    roughInterval = maxValue / 6 ' 6本程度の目盛り
    
    ' 切りの良い数字に調整
    Select Case True
        Case roughInterval <= 1: CalculateTickInterval = 1
        Case roughInterval <= 2: CalculateTickInterval = 2
        Case roughInterval <= 5: CalculateTickInterval = 5
        Case roughInterval <= 10: CalculateTickInterval = 10
        Case roughInterval <= 20: CalculateTickInterval = 20
        Case roughInterval <= 25: CalculateTickInterval = 25
        Case roughInterval <= 50: CalculateTickInterval = 50
        Case roughInterval <= 100: CalculateTickInterval = 100
        Case Else
            Dim magnitude As Long
            magnitude = Int(Log(roughInterval) / Log(10))
            Dim base As Double
            base = 10 ^ magnitude
            If roughInterval <= 2 * base Then
                CalculateTickInterval = 2 * base
            ElseIf roughInterval <= 5 * base Then
                CalculateTickInterval = 5 * base
            Else
                CalculateTickInterval = 10 * base
            End If
    End Select
End Function

Private Function GenerateCommentText(ByVal occurrenceValue As String, _
                                   ByVal dtStart As Date, _
                                   ByVal dtEnd As Date, _
                                   ByVal isProcessing As Boolean) As String
    ' コメントテキストの生成
    
    If isProcessing Then
        GenerateCommentText = "発生が「加工」のため、グラフは表示されません。"
    Else
        Dim startDateStr As String, endDateStr As String
        startDateStr = Format(dtStart, "m/d")
        endDateStr = Format(dtEnd, "m/d")
        GenerateCommentText = occurrenceValue & " 流出不良集計 " & _
                            startDateStr & " ～ " & endDateStr
    End If
End Function

Private Sub SetModeValidation(ByVal ws As Worksheet)
    ' モードフィールドの入力規則設定（高速化版）
    
    Dim modeItems As Object
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    ' 除外リスト
    Dim excludeItems As Object
    Set excludeItems = CreateObject("Scripting.Dictionary")
    excludeItems("A") = True
    excludeItems("B") = True
    excludeItems("C") = True
    excludeItems("D") = True
    excludeItems("E") = True
    excludeItems("Fr RH") = True
    
    ' AG列のデータを配列で一括取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        Dim dataArray As Variant
        dataArray = ws.Range("AG13:AG" & lastRow).Value
        
        Dim i As Long
        For i = 1 To UBound(dataArray, 1)
            Dim cellValue As String
            cellValue = Trim(CStr(dataArray(i, 1)))
            
            If cellValue <> "" And Not excludeItems.Exists(cellValue) Then
                modeItems(cellValue) = True
            End If
        Next i
    End If
    
    ' 入力規則の設定
    With ws.Range("T3")
        .Validation.Delete
        If modeItems.Count > 0 Then
            Dim modeList As String
            modeList = Join(modeItems.Keys, ",")
            .Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=modeList
            .Value = ""
        Else
            .Value = "モード項目なし"
        End If
    End With
    
    Set modeItems = Nothing
    Set excludeItems = Nothing
End Sub