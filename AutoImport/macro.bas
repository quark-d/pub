' 指定したモジュールのみをエクスポートするマクロ
Public Sub ExportSpecificModule(moduleName As String)
    Dim vbComp As VBIDE.VBComponent
    Dim exportFolder As String
    
    ' エクスポートフォルダのパスを設定
    exportFolder = "C:\path_to_export_folder\"
    
    ' 指定されたモジュールのみをエクスポート
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not vbComp Is Nothing Then
        vbComp.Export exportFolder & vbComp.Name & ".bas"
    Else
        MsgBox "指定したモジュールが見つかりません。"
    End If
End Sub
