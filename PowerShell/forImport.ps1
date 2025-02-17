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
$excel.Visible = $true  # Excelを表示

# Excelファイルを開く
$workbook = $excel.Workbooks.Open($excelFilePath)

# VBE (VBA Editor) の取得
$vbe = $excel.VBE
$vbProject = $workbook.VBProject
$modules = $vbProject.VBComponents

# 既存の同名モジュールがあれば削除
$existingModule = $null

foreach ($module in $modules) {
    # `InvokeMember` を使って COM オブジェクトの `Name` プロパティを取得
    $moduleNameFromExcel = $module.GetType().InvokeMember("Name", "GetProperty", $null, $module, $null)

    if ($moduleNameFromExcel -eq $moduleName) {
        $existingModule = $module
        break
    }
}

if ($existingModule -ne $null) {
    Write-Host "既存のモジュール [$moduleName] を削除します。"
    $modules.Remove($existingModule)
}

# モジュールをインポート
$modules.Import($vbaModulePath)
Write-Host "モジュール [$moduleName] をインポートしました。"

# 保存して閉じる
$workbook.Save()
$workbook.Close($false)

# Excelを終了
$excel.Quit()

# COMオブジェクトの解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($modules) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbProject) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbe) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "マクロのインポートが完了しました: $vbaModulePath → $excelFilePath"
