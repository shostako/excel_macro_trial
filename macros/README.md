# macros/ ディレクトリ

## 概要
動作確認済み・本番運用可能なVBAマクロライブラリ

## ファイル形式
- **エンコーディング**: UTF-8
- **状態**: テスト完了、本番利用可能

## ファイル命名規則
- **接頭辞**: `m` (モジュール識別)
- **形式**: `m処理名_対象別.bas`
- **例**: `m日別集計_品番別.bas`

## 使用方法
1. bas2sjis でShift-JIS変換
2. ExcelのVBEでインポート
3. そのまま本番利用可能