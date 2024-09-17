' 指定したモジュールを削除し、再インポートするマクロ
Public Sub ReplaceModule(moduleName As String, modulePath As String)
    Dim vbComp As VBIDE.VBComponent
    On Error Resume Next ' モジュールが存在しない場合に備える

    ' 指定したモジュールが存在する場合は削除
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
    
    ' 指定されたファイルをインポート
    ThisWorkbook.VBProject.VBComponents.Import modulePath
    On Error GoTo 0
End Sub
