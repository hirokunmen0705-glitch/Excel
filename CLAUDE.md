# Excel（業務ログ集計表 スクリプト群）

## 概要
業務ログ集計Excelの生成・管理スクリプト群。
`excel-automation/` リポジトリと内容が重複している場合は `excel-automation/` が最新版。

## 関連リポジトリ
- `excel-automation/` — アクティブな開発はこちら
- `gyomu-log/` — Excelファイル本体

## 起動前の必須手順
```bash
# Excelを必ず閉じてから実行（WSLからはpowershell.exe経由）
powershell.exe -Command "taskkill /F /IM EXCEL.EXE"

python3 rebuild_excel.py
```

## 既知の罠（Excel×VBA×openpyxl）
- `ReDim array(1 To 0)` → エラー9。`If count = 0 Then Exit Sub` ガード必須
- openpyxl生成xlsxはcalcChain.xmlを zipfile で除去してから保存
- Excelが月文字列を日付シリアル値に変換 → `TEXT(cell,"yyyy年mm月")` でラップ
- ゾンビExcel: テスト冒頭で `powershell.exe -Command "taskkill /F /IM EXCEL.EXE"`
- PowerShell COM VBA注入は2ステップ: ①xlsx→xlsm保存 ②再オープン→注入

## ファイル保存場所
- 完成版: `/mnt/c/Users/h-kaw/OneDrive/デスクトップ/`（ユーザーへ渡すファイルはここ）
- WSL側作業: `~/` 配下

## 開発ルール
- 常に日本語でコミュニケーション
- コミットは `/commit` スキルで行う
