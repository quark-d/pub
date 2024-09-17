' 指定したモジュールのみをインポートするマクロ
Public Sub ImportSpecificModule(moduleName As String)
    Dim vbComp As VBIDE.VBComponent
    Dim importFolder As String
    Dim filePath As String
    
    ' インポートフォルダのパスを設定
    importFolder = "C:\path_to_export_folder\"
    filePath = importFolder & moduleName & ".bas"
    
    ' モジュールが存在する場合、インポートする
    If Dir(filePath) <> "" Then
        ThisWorkbook.VBProject.VBComponents.Import filePath
    Else
        MsgBox "指定したモジュールのファイルが見つかりません。"
    End If
End Sub
