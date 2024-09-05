Private filterManager As FilterManager

Private Sub UserForm_Initialize()
    Dim headers As Variant
    headers = Array("ID", "名前", "住所") ' リストボックスの列名
    
    Set filterManager = New FilterManager
    filterManager.Initialize Me.ListBox1, headers ' リストボックスと列名を設定
    
    ' コンボボックスと対応する列名を追加
    filterManager.AddFilter Me.ComboBox1, "名前"
    filterManager.AddFilter Me.ComboBox2, "住所"
    
    ' コンボボックスにユニークな値を設定
    filterManager.PopulateComboBoxes
End Sub

Private Sub ComboBox1_Change()
    filterManager.ApplyFilters
End Sub

Private Sub ComboBox2_Change()
    filterManager.ApplyFilters
End Sub
