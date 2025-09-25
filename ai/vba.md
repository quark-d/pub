はい、どちらの方法も実装できます。

1. **VBA内にPowerShellスクリプトを“文字列として保持”**し、


2. 実行時にダウンロードフォルダへ .ps1 として書き出して実行、がシンプルで運用もしやすいです。



下に“そのまま動く最小実装”を置きます。日本語を含むスクリプトでも文字化けしないよう**UTF-8(BOM付き)**で保存しています。


---

手順（VBA）

1) 標準モジュールに貼り付け

Option Explicit

' ▼ PowerShellスクリプト本文をここに保持（VBAの定数文字列）
'   ダブルクォーテーションは "" と二重にエスケープします。
Private Const PS_SCRIPT As String = _
    "param(" & vbCrLf & _
    "  [Parameter(Mandatory=$true)]" & vbCrLf & _
    "  [string]$FolderPath," & vbCrLf & _
    "  [Parameter(Mandatory=$true)]" & vbCrLf & _
    "  [string]$Url," & vbCrLf & _
    "  [ValidateSet('edge','default')]" & vbCrLf & _
    "  [string]$BrowserMode = 'edge'" & vbCrLf & _
    ")" & vbCrLf & _
    "" & vbCrLf & _
    "Add-Type -Namespace Win32 -Name User32 -MemberDefinition @'" & vbCrLf & _
    "using System;" & vbCrLf & _
    "using System.Runtime.InteropServices;" & vbCrLf & _
    "public class Native {" & vbCrLf & _
    "  [DllImport(""user32.dll"", SetLastError=true)]" & vbCrLf & _
    "  public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);" & vbCrLf & _
    "}" & vbCrLf & _
    "'@" & vbCrLf & _
    "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
    "$work = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea" & vbCrLf & _
    "function Move-ToRect([IntPtr]$h, [int]$x, [int]$y, [int]$w, [int]$hgt) {" & vbCrLf & _
    "  if ($h -eq [IntPtr]::Zero) { throw 'HWND not found' }" & vbCrLf & _
    "  [Win32.User32+Native]::MoveWindow($h, $x, $y, $w, $hgt, $true) | Out-Null" & vbCrLf & _
    "}" & vbCrLf & _
    "$null = Start-Process explorer.exe -ArgumentList ""`"$FolderPath`"""" & vbCrLf & _
    "$shell = New-Object -ComObject Shell.Application" & vbCrLf & _
    "function Get-ExplorerHwndForPath($target) {" & vbCrLf & _
    "  $target = (Resolve-Path $target).Path.TrimEnd('\')" & vbCrLf & _
    "  for ($i=0; $i -lt 50; $i++) {" & vbCrLf & _
    "    foreach ($w in $shell.Windows()) {" & vbCrLf & _
    "      try { if ($w.FullName -like '*explorer.exe') {" & vbCrLf & _
    "        $loc = [uri]$w.LocationURL" & vbCrLf & _
    "        $p = [System.Web.HttpUtility]::UrlDecode($loc.AbsolutePath).Replace('/','\')" & vbCrLf & _
    "        if ($loc.Scheme -eq 'file' -and $p) {" & vbCrLf & _
    "          if ($p -match '^[A-Za-z]:\\') { $p = $p.TrimEnd('\') }" & vbCrLf & _
    "          if ($p -notmatch '^[A-Za-z]:\\' -and $p -notmatch '^\\\\\\\\') { $p = '\\\\' + $p }" & vbCrLf & _
    "          if ($p -ieq $target) { return [IntPtr]$w.HWND }" & vbCrLf & _
    "        } } } catch {}" & vbCrLf & _
    "    } Start-Sleep -Milliseconds 120" & vbCrLf & _
    "  } return [IntPtr]::Zero" & vbCrLf & _
    "}" & vbCrLf & _
    "$explorerHwnd = Get-ExplorerHwndForPath $FolderPath" & vbCrLf & _
    "" & vbCrLf & _
    "[IntPtr]$browserHwnd = [IntPtr]::Zero" & vbCrLf & _
    "if ($BrowserMode -eq 'edge') {" & vbCrLf & _
    "  $edge = Start-Process msedge.exe -ArgumentList ""--new-window `"$Url`"""" -PassThru" & vbCrLf & _
    "  for ($i=0; $i -lt 50 -and $browserHwnd -eq [IntPtr]::Zero; $i++) {" & vbCrLf & _
    "    try { $edge.Refresh() | Out-Null; if ($edge.MainWindowHandle -ne 0) { $browserHwnd = [IntPtr]$edge.MainWindowHandle } } catch {}" & vbCrLf & _
    "    Start-Sleep -Milliseconds 120" & vbCrLf & _
    "  }" & vbCrLf & _
    "} else {" & vbCrLf & _
    "  Start-Process cmd.exe -ArgumentList ""/c start `"$Url`"""" -WindowStyle Hidden" & vbCrLf & _
    "  Start-Sleep -Seconds 2" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "$halfW = [int]([math]::Floor($work.Width / 2))" & vbCrLf & _
    "$leftRect  = @{ x=$work.X; y=$work.Y; w=$halfW; h=$work.Height }" & vbCrLf & _
    "$rightRect = @{ x=$work.X + $halfW; y=$work.Y; w=$work.Width - $halfW; h=$work.Height }" & vbCrLf & _
    "if ($explorerHwnd -ne [IntPtr]::Zero) { Move-ToRect $explorerHwnd $leftRect.x $leftRect.y $leftRect.w $leftRect.h }" & vbCrLf & _
    "if ($browserHwnd  -ne [IntPtr]::Zero) { Move-ToRect $browserHwnd  $rightRect.x  $rightRect.y  $rightRect.w  $rightRect.h }"

' ▼ エントリーポイント例
Public Sub RunExplorerAndSharePoint()
    Dim folder As String: folder = "Z:\Team\Shared" ' ネットワークドライブ or UNC
    Dim url As String:    url    = "https://contoso.sharepoint.com/sites/Team"
    Dim browserMode As String: browserMode = "edge" ' or "default"

    Dim ps1Path As String
    ps1Path = GetDownloadsPath() & "\OpenExplAndBrowser.ps1"

    ' 1) PSスクリプトをUTF-8(BOM)で保存（上書きOK）
    WriteTextUtf8BOM ps1Path, PS_SCRIPT

    ' 2) PowerShellを呼び出し（プロセス限定でExecutionPolicyをBypass）
    Dim psExe As String
    psExe = GetPowerShellExePath()  ' Windows PowerShell / PowerShell(Core) どちらでも

    Dim cmd As String
    cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -File """ & ps1Path & _
          """ -FolderPath """ & folder & """ -Url """ & url & """ -BrowserMode " & browserMode

    CreateObject("WScript.Shell").Run cmd, 0, False
End Sub

' ▼ ダウンロードフォルダのパスを取得（一般的には USERPROFILE\Downloads でOK）
Private Function GetDownloadsPath() As String
    GetDownloadsPath = Environ$("USERPROFILE") & "\Downloads"
End Function

' ▼ 利用可能な PowerShell 実行ファイルのパスを返す
Private Function GetPowerShellExePath() As String
    Dim p As String
    ' 1) PowerShell(Core) (pwsh.exe) 優先
    p = Environ$("ProgramFiles") & "\PowerShell\7\pwsh.exe"
    If Dir$(p) <> "" Then GetPowerShellExePath = p: Exit Function
    ' 2) Windows PowerShell
    p = Environ$("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    If Dir$(p) <> "" Then GetPowerShellExePath = p: Exit Function
    ' フォールバック
    GetPowerShellExePath = "powershell.exe"
End Function

' ▼ UTF-8(BOM) でファイルに書き出す（ADODB.Stream 使用）
Private Sub WriteTextUtf8BOM(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                   ' text
    stm.Charset = "utf-8"
    stm.Open
    ' BOMを明示書き込み：UTF-8のBOMシーケンス EF BB BF
    Dim bom() As Byte
    bom = StrConv(ChrW(&HEF) & ChrW(&HBB) & ChrW(&HBF), vbFromUnicode)
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1: bin.Open
    bin.Write bom
    ' 本文
    stm.WriteText text
    stm.Position = 0
    bin.Position = bin.Size
    stm.CopyTo bin
    bin.SaveToFile path, 2         ' adSaveCreateOverWrite
    bin.Close: stm.Close
End Sub


---

仕組みと注意点

VBA側でハードコード → 実行時に %USERPROFILE%\Downloads に .ps1 を出力 → ExecutionPolicyは“プロセス限定”でBypassして起動。
会社規定が厳しい場合は、Bypassの扱いを管理部門と合意ください（恒久設定は不要です）。

文字コード：日本語を含むため UTF-8(BOM) 推奨。上記 WriteTextUtf8BOM はBOM付きで保存します。
（CreateTextFile(..., , True) のUnicode保存でも動きますが UTF-16LE になるため、他ツール連携を考えるとUTF-8が無難）

PowerShell実行ファイル：

PowerShell 7（pwsh.exe）があれば優先、なければ Windows PowerShell（powershell.exe）を使用。


パスに空白があっても全て二重引用符で囲っているので安全です。

更新運用：VBAの PS_SCRIPT を修正 → 次回実行時に同名 .ps1 を上書きします。



---

代替案（必要なら）

一時ファイルに保存：Environ$("TEMP") & "\xxxxx.ps1" に書き出し、実行後に削除。

Base64埋め込み：VBAにはBase64化したスクリプトを埋め込み、実行前にデコード→保存。長文を1行に収めたい時に有効。

保存せず直実行：-Command 引数でインライン実行も可能ですが、今回のような数百行規模は可読性とデバッグ性が低下するので非推奨。



---

必要であれば、上の PS_SCRIPT に以前お渡ししたフル機能版（Explorerとブラウザを左右整列）をそのまま差し替えた完全版も作ってお渡しします。ブラウザ固定（Edge/Chrome）やマルチモニタ対応への拡張もすぐ対応できます。

