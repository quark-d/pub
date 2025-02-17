# RunMacroExportAll.ps1

param (
    [string]$excelFilePath,   # Excelファイルのパス
    [string]$exportFolder     # エクスポート先フォルダ
)

Write-Host "Excelファイル: $excelFilePath"
Write-Host "エクスポート先フォルダ: $exportFolder"

# エクスポートフォルダの作成（存在しない場合）
if (!(Test-Path -Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder | Out-Null
    Write-Host "エクスポートフォルダを作成しました: $exportFolder"
}

# Excelの起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Excelを非表示で実行

# Excelファイルを開く
$workbook = $excel.Workbooks.Open($excelFilePath)

# VBE (VBA Editor) の取得
$vbe = $excel.VBE
$vbProject = $workbook.VBProject

# VBAプロジェクトのモジュール一覧を取得
$modules = $vbProject.VBComponents

foreach ($module in $modules) {
    $moduleName = $module.Name
    $moduleType = $module.Type

    # エクスポートファイルの拡張子を決定
    $extension = ".bas"  # 通常のモジュール
    if ($moduleType -eq 2) { $extension = ".cls" }  # クラスモジュール
    if ($moduleType -eq 3) { $extension = ".frm" }  # ユーザーフォーム

    $exportPath = "$exportFolder\$moduleName$extension"

    # モジュールをエクスポート
    $module.Export($exportPath)
    Write-Host "エクスポート完了: $exportPath"
}

# Excelを閉じる
$workbook.Close($false)
$excel.Quit()

# COMオブジェクトの解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($modules) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbProject) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbe) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "すべてのVBAモジュールをエクスポートしました。"
