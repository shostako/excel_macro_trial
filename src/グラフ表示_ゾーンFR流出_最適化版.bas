Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Option Explicit

Sub グラフ表示_ゾーンFR流出()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ
    ' 作成日: 2025/06/03
    ' 更新日: 2025/06/04 (グラフ軸動的調整機能追加・改良版v4 - 目盛り間隔も動的調整)
    ' 更新日: 2025/06/06 (モードフィールド入力規則とフィルタ機能追加)
    ' 更新日: 2025/06/07 (最適化版 - 処理速度改善とステータスバー詳細化)

    Dim ws As Worksheet
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String ' E3: 発生
    Dim discovery2Value As String ' E4: 発見2
    Dim dictDiscovery2 As Object  ' 発見2の高速検索用Dictionary
    Dim isProcessing As Boolean   ' 「発生」が「加工」工程判定用
    Dim isMould As Boolean        ' 「発生」が「モール」工程判定用
    Dim isDiscovery2Empty As Boolean ' 発見2の値が空か判定用
    Dim commentText As String     ' D6に設定するコメント用
    Dim ptArray As Variant        ' ピボットテーブル配列
    Dim i As Long

    ' エラー処理を設定
    On Error GoTo ErrorHandler

    ' 三種の神器 - 最重要の最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "初期化中..."

    ' ワークシートの取得（Activateは使わない）
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

    ' 「発生」が「加工」かどうかを判定
    isProcessing = (occurrenceValue = "加工")

    ' 「発生」が「モール」かどうかを判定
    isMould = (occurrenceValue = "モール")

    ' 発見2値をDictionaryに格納（高速検索用）
    Set dictDiscovery2 = CreateObject("Scripting.Dictionary")
    If Not isDiscovery2Empty Then
        Dim arrDiscovery2 As Variant
        arrDiscovery2 = Split(discovery2Value, ",")
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            dictDiscovery2(Trim(arrDiscovery2(i))) = True
        Next i
    End If

    ' 全ピボットテーブルの手動更新モードに設定（一括処理の準備）
    Application.StatusBar = "ピボットテーブル更新準備中..."
    ptArray = Array(pt1, pt2, pt3, pt4, pt5)
    For i = 0 To UBound(ptArray)
        ptArray(i).ManualUpdate = True
    Next i

    ' モード2フィルタをリセット（全て表示に戻す）
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(ptArray)

    ' 各ピボットテーブルのフィルタ設定（個別にステータス表示）
    Application.StatusBar = "アルヴェル Fr ピボットテーブルを設定中..."
    Call ApplyPivotFilter(pt1, dtStart, dtEnd, "アルヴェル", "Fr", occurrenceValue, dictDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "アルヴェル Rr ピボットテーブルを設定中..."
    Call ApplyPivotFilter(pt2, dtStart, dtEnd, "アルヴェル", "Rr", occurrenceValue, dictDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Fr ピボットテーブルを設定中..."
    Call ApplyPivotFilter(pt3, dtStart, dtEnd, "ノアヴォク", "Fr", occurrenceValue, dictDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Rr ピボットテーブルを設定中..."
    Call ApplyPivotFilter(pt4, dtStart, dtEnd, "ノアヴォク", "Rr", occurrenceValue, dictDiscovery2, isDiscovery2Empty)
    
    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call ApplyPivotFilterForMode(pt5, dtStart, dtEnd, occurrenceValue, dictDiscovery2, isDiscovery2Empty)

    ' ピボットテーブルの一括更新実行
    Application.StatusBar = "ピボットテーブルを更新中..."
    For i = 0 To UBound(ptArray)
        ptArray(i).ManualUpdate = False
        ptArray(i).RefreshTable
    Next i

    ' グラフ表示設定の決定
    Application.StatusBar = "グラフ表示設定を処理中..."
    Dim showGraph1 As Boolean, showGraph2 As Boolean
    Dim showGraph3 As Boolean, showGraph4 As Boolean
    Dim startDateStr As String, endDateStr As String

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
            showGraph1 = True  ' グラフ1: 表示
            showGraph2 = True  ' グラフ2: 表示
            showGraph3 = False ' グラフ3: 非表示
            showGraph4 = False ' グラフ4: 非表示
            
            ' 日付を M/D 形式に変換
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
            
        Case Else
            ' その他の場合（全てのグラフを表示）
            showGraph1 = True
            showGraph2 = True
            showGraph3 = True
            showGraph4 = True
            
            ' 日付を M/D 形式に変換
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    End Select

    ' グラフ表示/非表示の一括適用
    Application.StatusBar = "グラフの表示設定を適用中..."
    Call SetChartVisibilityBatch(ws, Array("グラフ1", "グラフ2", "グラフ3", "グラフ4"), _
                                 Array(showGraph1, showGraph2, showGraph3, showGraph4))

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
    axisMax = GetNiceMaxValue(overallMax)
    tickInterval = GetNiceTickInterval(axisMax)
    
    ' 表示されているグラフに軸設定を適用
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
    Call SetupModeFieldValidation(ws)

Cleanup:
    Application.StatusBar = "処理が完了しました。"
    Application.Wait Now + TimeValue("00:00:01") ' 1秒間表示
    Application.StatusBar = False ' ステータスバーをクリア
    
    ' 三種の神器を元に戻す
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    ' オブジェクトの解放
    Set dictDiscovery2 = Nothing
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

Private Sub ApplyPivotFilter(ByVal pt As PivotTable, _
                            ByVal startDate As Date, _
                            ByVal endDate As Date, _
                            ByVal alNoahFilter As String, _
                            ByVal frRrFilter As String, _
                            ByVal occurrenceFilter As String, _
                            ByVal dictDiscovery2 As Object, _
                            ByVal isDiscovery2Empty As Boolean)
    ' ピボットテーブルのフィルタリング（最適化版）
    
    Dim pi As PivotItem
    Dim d As Date
    
    On Error Resume Next
    
    ' 日付フィールドのフィルタリング（一括処理）
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
    
    ' ページフィールドの設定
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールドのフィルタリング（Dictionary使用で高速化）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = dictDiscovery2.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub ApplyPivotFilterForMode(ByVal pt As PivotTable, _
                                   ByVal startDate As Date, _
                                   ByVal endDate As Date, _
                                   ByVal occurrenceFilter As String, _
                                   ByVal dictDiscovery2 As Object, _
                                   ByVal isDiscovery2Empty As Boolean)
    ' モード抽出用ピボットテーブルのフィルタリング
    
    Dim pi As PivotItem
    Dim d As Date
    
    On Error Resume Next
    
    ' 日付フィールドのフィルタリング
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
    
    ' アル/ノア・Fr/Rr：全表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter
    
    ' 発見2フィールド（Dictionary使用）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = dictDiscovery2.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub ResetMode2Filters(ByVal ptArray As Variant)
    ' モード2フィルタの一括リセット
    Dim pt As PivotTable
    Dim i As Long
    
    On Error Resume Next
    For i = 0 To UBound(ptArray)
        Set pt = ptArray(i)
        With pt.PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = "(すべて)"
        End With
    Next i
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, _
                                   ByVal chartNames As Variant, _
                                   ByVal visibilities As Variant)
    ' グラフの表示/非表示を一括設定
    Dim i As Long
    Dim chObj As ChartObject
    
    On Error Resume Next
    For i = 0 To UBound(chartNames)
        Set chObj = ws.ChartObjects(chartNames(i))
        If Not chObj Is Nothing Then
            chObj.Visible = visibilities(i)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    ' ピボットテーブルの最大値を高速取得（配列使用）
    Dim dataRange As Range
    Dim arr As Variant
    Dim maxVal As Double
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    Set dataRange = pt.DataBodyRange
    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' データを配列に一括読み込み
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
    ' データの最大値から「良い感じの」軸の最大値を計算
    
    If maxValue <= 0 Then
        GetNiceMaxValue = 10
        Exit Function
    End If
    
    ' 目標：最大値の110%～120%で最も近い切りの良い数字
    Dim targetMin As Double, targetMax As Double
    targetMin = maxValue * 1.1
    targetMax = maxValue * 1.2
    
    ' 桁数を取得
    Dim magnitude As Long
    magnitude = Int(Log(targetMin) / Log(10))
    
    ' 基本単位
    Dim base As Double
    base = 10 ^ magnitude
    
    ' 候補値を試す
    Dim candidates As Variant
    candidates = Array(1, 1.5, 2, 2.5, 3, 4, 5, 6, 7, 8, 9, 10)
    
    Dim i As Long
    For i = 0 To UBound(candidates)
        Dim candidate As Double
        candidate = candidates(i) * base
        If candidate >= targetMin And candidate <= targetMax Then
            GetNiceMaxValue = candidate
            Exit Function
        End If
    Next i
    
    ' 適切な値が見つからない場合
    GetNiceMaxValue = targetMax
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    ' 軸の最大値に基づいて適切な目盛り間隔を計算
    
    ' 理想的な目盛り数は6～8本
    Dim targetTicks As Long
    targetTicks = 6
    
    Dim roughInterval As Double
    roughInterval = maxValue / targetTicks
    
    ' 桁数を取得
    Dim magnitude As Long
    magnitude = Int(Log(roughInterval) / Log(10))
    
    ' 基本単位
    Dim base As Double
    base = 10 ^ magnitude
    
    ' 切りの良い間隔に調整
    Dim normalized As Double
    normalized = roughInterval / base
    
    If normalized <= 1 Then
        GetNiceTickInterval = base
    ElseIf normalized <= 2 Then
        GetNiceTickInterval = 2 * base
    ElseIf normalized <= 5 Then
        GetNiceTickInterval = 5 * base
    Else
        GetNiceTickInterval = 10 * base
    End If
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, _
                                ByVal chartName As String, _
                                ByVal maxValue As Double, _
                                ByVal tickInterval As Double)
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

Private Sub SetupModeFieldValidation(ByVal ws As Worksheet)
    ' モードフィールドの入力規則設定
    Dim modeDict As Object
    Dim excludeDict As Object
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim cellValue As String
    Dim modeList As String
    
    ' Dictionary初期化
    Set modeDict = CreateObject("Scripting.Dictionary")
    Set excludeDict = CreateObject("Scripting.Dictionary")
    
    ' 除外する値を設定
    excludeDict("A") = True
    excludeDict("B") = True
    excludeDict("C") = True
    excludeDict("D") = True
    excludeDict("E") = True
    excludeDict("Fr RH") = True
    
    On Error Resume Next
    
    ' AG列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        ' AG13以降のデータを配列で処理（高速化）
        Set rng = ws.Range("AG13:AG" & lastRow)
        Dim arr As Variant
        arr = rng.Value
        
        Dim i As Long
        For i = 1 To UBound(arr, 1)
            cellValue = Trim(CStr(arr(i, 1)))
            If cellValue <> "" And Not excludeDict.Exists(cellValue) And Not modeDict.Exists(cellValue) Then
                modeDict(cellValue) = True
            End If
        Next i
    End If
    
    ' リスト文字列作成と入力規則設定
    If modeDict.Count > 0 Then
        modeList = Join(modeDict.Keys, ",")
        
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
    
    ' オブジェクトの解放
    Set modeDict = Nothing
    Set excludeDict = Nothing
End Sub