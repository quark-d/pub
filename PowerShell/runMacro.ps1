# コマンドライン引数からExcelファイルのパスを受け取る
param (
    [string]$excelFilePath
)

# 実行するマクロ名の指定（例としてYourMacroNameとしています）
$macroName = "YourMacroName"

# Excelアプリケーションの起動
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false

# ワークブックの開
$Workbook = $Excel.Workbooks.Open($excelFilePath)

# マクロの実行
$Excel.Run($macroName)

# ワークブックの閉じる
$Workbook.Close($false)

# Excelアプリケーションの終了
$Excel.Quit()
