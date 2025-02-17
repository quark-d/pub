# RunMacro.ps1

param (
    [string]$excelFilePath,  # Excelファイルのパス
    [string]$vbaModulePath,  # インポートするVBAモジュールのパス
    [string]$moduleName      # モジュール名（拡張子なし）
)

Write-Host "Excelファイル: $excelFilePath"
Write-Host "VBAモジュール: $vbaModulePath"
Write-Host "モジュール名: $moduleName"

# Excelの起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true  # Excelを表示（デバッグ用）

# Excelファイルを開く
$workbook = $excel.Workbooks.Open($excelFilePath)

# VBAマクロを実行（モジュール削除 → インポート）
$excel.Run("ImportModule", $moduleName, $vbaModulePath)

# 保存して閉じる
$workbook.Save()
$workbook.Close($false)

# Excelを終了
$excel.Quit()

# COMオブジェクトの解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "マクロの実行が完了しました。"
