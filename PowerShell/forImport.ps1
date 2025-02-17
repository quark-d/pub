# RunMacro.ps1

param (
    [string]$excelFilePath,
    [string]$vbaModulePath
)

# Excelの起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true  # Excelを表示

# Excelファイルを開く
$workbook = $excel.Workbooks.Open($excelFilePath)

# VBE (VBA Editor) の取得
$vbe = $excel.VBE

# 現在のVBAプロジェクトのモジュール一覧を取得
$vbProject = $workbook.VBProject
$modules = $vbProject.VBComponents

# 既存の同名モジュールがあれば削除
foreach ($module in $modules) {
    if ($module.Name -eq "Module1") {
        $modules.Remove($module)
        break
    }
}

# モジュールをインポート
$modules.Import($vbaModulePath)

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
