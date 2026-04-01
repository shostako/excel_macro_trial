Attribute VB_Name = "m転記_段取加工"
Option Explicit

' ========================================
' マクロ名: 転記_段取加工
' 処理概要: 段取加工テーブルから集計表T列へ段取データを転記
' ソーステーブル: テーブル「_段取加工」（シート不問、テーブル名で検索）
' 基準日付: シート「集計表」セルA1の日付
' 転記先: シート「集計表」T列（行37-42）
' 転記指標: 6種類（日平均時間、日段取時間、日段取回数、累計段取時間、累計段取回数、平均段取時間）
' 時間換算: なし（元データが分単位のためそのまま転記）
' ========================================
Sub 転記_段取加工()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Application.StatusBar = "段取加工データの転記を開始します..."

    ' =================================
    ' 第1段階：シートとテーブルの取得・検証
    ' =================================

    ' 転記先シート（集計表）の取得
    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler

    ' 基準日付の取得（集計表A1セルから）
    Dim targetDate As Date
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。", vbCritical, "日付エラー"
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value

    ' ソーステーブル（_段取加工）をブック全体から検索
    Dim sourceTable As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "_段取加工" Then
                Set sourceTable = lo
                Exit For
            End If
        Next lo
        If Not sourceTable Is Nothing Then Exit For
    Next ws

    If sourceTable Is Nothing Then
        MsgBox "テーブル「_段取加工」が見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If

    ' データ範囲の確認
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_段取加工」テーブルにデータがありません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    Dim sourceData As Range
    Set sourceData = sourceTable.DataBodyRange

    ' =================================
    ' 第2段階：基準日付に一致する行の検索
    ' =================================

    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_段取加工」テーブルに「日付」列が見つかりません。", vbCritical, "列エラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler

    Dim sourceRow As Long
    sourceRow = 0
    Dim j As Long
    For j = 1 To sourceData.Rows.Count
        If sourceData.Cells(j, dateColIndex).Value = targetDate Then
            sourceRow = j
            Exit For
        End If
    Next j

    If sourceRow = 0 Then
        MsgBox "日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If

    ' =================================
    ' 第3段階：段取データの転記処理
    ' =================================

    Application.StatusBar = "段取加工データ転記中..."

    ' 転記マッピング定義（ソース列名, 転記先行, 転記先列）
    ' T列 = 20列目、行37-42
    Dim transfers() As Variant
    transfers = Array( _
        Array("日平均時間", 37, 20), _
        Array("日段取時間", 38, 20), _
        Array("日段取回数", 39, 20), _
        Array("累計段取時間", 40, 20), _
        Array("累計段取回数", 41, 20), _
        Array("平均段取時間", 42, 20) _
    )

    Dim i As Long
    Dim transferItem As Variant
    Dim columnName As String
    Dim colIndex As Long
    Dim sourceValue As Variant
    Dim targetRow As Long
    Dim targetCol As Long

    For i = 0 To UBound(transfers)
        transferItem = transfers(i)
        columnName = transferItem(0)
        targetRow = transferItem(1)
        targetCol = transferItem(2)

        On Error Resume Next
        colIndex = sourceTable.ListColumns(columnName).Index

        If Err.Number = 0 Then
            sourceValue = sourceData.Cells(sourceRow, colIndex).Value

            ' 空白・NULL値の処理
            If IsEmpty(sourceValue) Or sourceValue = "" Or IsNull(sourceValue) Then
                sourceValue = 0
            End If

            ' 時間関連：そのまま転記（元データは分単位）+ 書式設定
            If InStr(columnName, "時間") > 0 Then
                wsTarget.Cells(targetRow, targetCol).Value = sourceValue
                ' 平均系は小数点1桁、その他は整数
                If InStr(columnName, "平均") > 0 Then
                    wsTarget.Cells(targetRow, targetCol).NumberFormatLocal = "_-* 0.0"" 分"""
                Else
                    wsTarget.Cells(targetRow, targetCol).NumberFormatLocal = "_-* 0"" 分"""
                End If
            Else
                ' 回数関連：そのまま転記
                wsTarget.Cells(targetRow, targetCol).Value = sourceValue
                wsTarget.Cells(targetRow, targetCol).NumberFormatLocal = "_-* 0"" 回"""
            End If
        Else
            Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i

    ' 正常終了
    Application.StatusBar = "段取加工データの転記完了"
    Application.Wait Now + TimeValue("0:00:01")
    GoTo CleanupAndExit

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "転記エラー"

CleanupAndExit:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
End Sub
