' FilterManager.cls

Private listBox As MSForms.ListBox
Private comboBoxes As Collection
Private filterColumns As Collection
Private columnNames As Collection
Private originalData As Collection ' 元のデータを保存するコレクション

' コンボボックスとリストボックスの初期化
Public Sub Initialize(ByVal targetListBox As MSForms.ListBox, ByVal headers As Variant)
    Set listBox = targetListBox
    Set comboBoxes = New Collection
    Set filterColumns = New Collection
    Set columnNames = New Collection
    Set originalData = New Collection
    
    ' 列名をコレクションに保存
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        columnNames.Add headers(i)
    Next i
    
    ' リストボックスの元データを保存
    Call StoreOriginalData
End Sub

' リストボックスの元データを保存する関数
Private Sub StoreOriginalData()
    Dim i As Long, j As Long
    Dim rowData As Variant
    
    ' リストボックスのすべての行データを保存
    For i = 0 To listBox.ListCount - 1
        rowData = listBox.List(i)
        originalData.Add rowData
    Next i
End Sub

' 列名から列インデックスを取得する関数
Private Function GetColumnIndexByName(ByVal columnName As String) As Integer
    Dim i As Long
    For i = 1 To columnNames.Count
        If columnNames(i) = columnName Then
            GetColumnIndexByName = i - 1 ' 0ベースインデックスに変換
            Exit Function
        End If
    Next i
    ' 見つからない場合はエラー
    Err.Raise vbObjectError + 513, "FilterManager", "列名が見つかりません: " & columnName
End Function

' コンボボックスを追加し、列名でフィルタを設定
Public Sub AddFilter(ByVal comboBox As MSForms.ComboBox, ByVal columnName As String)
    comboBoxes.Add comboBox
    filterColumns.Add GetColumnIndexByName(columnName)
End Sub

' コンボボックスにユニークな値を設定
Public Sub PopulateComboBoxes()
    Dim i As Long
    Dim j As Long
    Dim uniqueItems As Collection
    Dim item As String
    Dim columnIndex As Integer
    Dim comboBox As MSForms.ComboBox
    
    ' 各コンボボックスに対応するユニークなフィルタを設定
    For j = 1 To comboBoxes.Count
        Set comboBox = comboBoxes(j)
        columnIndex = filterColumns(j)
        Set uniqueItems = New Collection
        
        ' リストボックスの各アイテムをチェックし、ユニークな値を追加
        On Error Resume Next
        For i = 0 To listBox.ListCount - 1
            item = listBox.List(i, columnIndex)
            uniqueItems.Add item, CStr(item)
        Next i
        On Error GoTo 0
        
        ' コンボボックスにユニークな値を設定
        comboBox.Clear
        For i = 1 To uniqueItems.Count
            comboBox.AddItem uniqueItems(i)
        Next i
    Next j
End Sub

' フィルタを適用してリストボックスを絞り込み
Public Sub ApplyFilters()
    Dim i As Long, j As Long
    Dim comboBox As MSForms.ComboBox
    Dim columnIndex As Integer
    Dim filterValue As String
    Dim filteredItems As Collection
    Dim matches As Boolean
    Dim rowData As Variant
    
    ' フィルタされたアイテムを格納するコレクション
    Set filteredItems = New Collection
    
    ' 元の全データから絞り込みを開始
    For i = 1 To originalData.Count
        rowData = originalData(i)
        matches = True
        
        ' 各コンボボックスの値に基づいてフィルタを適用
        For j = 1 To comboBoxes.Count
            Set comboBox = comboBoxes(j)
            columnIndex = filterColumns(j)
            filterValue = comboBox.Value
            
            ' コンボボックスに選択値がある場合のみフィルタを適用
            If filterValue <> "" And rowData(columnIndex) <> filterValue Then
                matches = False
                Exit For
            End If
        Next j
        
        ' フィルタ条件に一致するアイテムをコレクションに追加
        If matches Then
            filteredItems.Add rowData
        End If
    Next i
    
    ' リストボックスをクリア
    listBox.Clear
    
    ' フィルタされたアイテムをリストボックスに追加
    For i = 1 To filteredItems.Count
        rowData = filteredItems(i)
        listBox.AddItem rowData(0) ' リストボックスに追加
        For j = 1 To UBound(rowData)
            listBox.List(listBox.ListCount - 1, j) = rowData(j)
        Next j
    Next i
End Sub

' Private filterManager As FilterManager

' Private Sub UserForm_Initialize()
'     Dim headers As Variant
'     headers = Array("ID", "名前", "住所") ' リストボックスの列名
    
'     Set filterManager = New FilterManager
'     filterManager.Initialize Me.ListBox1, headers ' リストボックスと列名を設定
    
'     ' コンボボックスと対応する列名を追加
'     filterManager.AddFilter Me.ComboBox1, "名前"
'     filterManager.AddFilter Me.ComboBox2, "住所"
    
'     ' コンボボックスにユニークな値を設定
'     filterManager.PopulateComboBoxes
' End Sub

' Private Sub ComboBox1_Change()
'     filterManager.ApplyFilters
' End Sub

' Private Sub ComboBox2_Change()
'     filterManager.ApplyFilters
' End Sub


