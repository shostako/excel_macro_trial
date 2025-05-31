Attribute VB_Name = "m集計表データクリア_改良版"
Option Explicit

' 集計表データクリアマクロ（改良版）
' 複数の集計表シートを一括削除する
Sub 集計表データクリア()
    ' 高速化設定（最優先）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Long
    Dim deletedCount As Long
    Dim totalCount As Long
    
    On Error GoTo ErrorHandler
    
    ' ========== 基本設定 ==========
    Set wb = ThisWorkbook
    deletedCount = 0
    
    ' 削除対象シート名
    sheetNames = Array("日別集計_モールFR別", "集計表_TG作業者別", "集計表_TG品番別", _
                      "集計表_モールFR別", "集計表_加工作業者別", "集計表_加工品番別", _
                      "集計表_塗装品番別", "集計表_流出廃棄")
    
    totalCount = UBound(sheetNames) - LBound(sheetNames) + 1
    
    Application.StatusBar = "集計表シート削除処理を開始します..."
    
    ' ========== シート削除処理 ==========
    For i = LBound(sheetNames) To UBound(sheetNames)
        Application.StatusBar = "シート削除中... (" & (i + 1) & "/" & totalCount & ")"
        
        On Error Resume Next
        Set ws = wb.Sheets(sheetNames(i))
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
            End If
            
            Set ws = Nothing
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' ========== 処理完了 ==========
    Application.StatusBar = "完了: " & deletedCount & "個のシートを削除しました"
    
    MsgBox "集計表データをクリアしました。" & vbCrLf & _
           "削除されたシート数: " & deletedCount & "/" & totalCount, vbInformation
    
Cleanup:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    Set wb = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    GoTo Cleanup
End Sub