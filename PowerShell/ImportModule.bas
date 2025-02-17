Sub ImportModule(ByVal moduleName As String, ByVal modulePath As String)
    Dim vbComp As VBComponent
    Dim vbProj As VBProject

    ' VBAプロジェクトへのアクセスを許可（必要に応じてExcelのセキュリティ設定を確認）
    Set vbProj = ThisWorkbook.VBProject

    ' 既存のモジュールを削除（同じ名前のものがあれば）
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(moduleName)
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
    On Error GoTo 0

    ' 新しいモジュールをインポート
    vbProj.VBComponents.Import modulePath

    MsgBox "モジュール " & moduleName & " をインポートしました。", vbInformation, "完了"
End Sub
