Private Sub UserForm_Initialize()
    Dim wsConfig As Worksheet
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim itm As ListItem
    Dim i As Integer
    Dim colName As Variant
    Dim colIndex As Integer
    Dim selectedColumns As New Collection
    Dim colWidth As Integer
    Dim filterCriteria As String
    Dim matchRow As Boolean

    ' フィルタリングの条件を設定（例として "name" 列でのフィルタリング）
    filterCriteria = Me.ComboBox1.Value ' コンボボックスで選択された値をフィルタ条件とする
    
    ' 表示する列名をCollectionに追加
    selectedColumns.Add "ID"
    selectedColumns.Add "name"
    selectedColumns.Add "addr"
    selectedColumns.Add "addr_add"
    
    ' ワークシートとテーブルを指定
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    
    ' 設定シートの確認または作成
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("ConfigListView")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets.Add
        wsConfig.Name = "ConfigListView"
    End If

    ' ListViewの設定
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .MultiSelect = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        ' コレクション内の列名を元にヘッダを追加
        For Each colName In selectedColumns
            colIndex = tbl.ListColumns(colName).index
            .ColumnHeaders.Add , , colName, 100 ' 固定幅を設定

            ' 保存されている幅があれば適用
            colWidth = wsConfig.Cells(1, colIndex).value
            If colWidth > 0 Then
                .ColumnHeaders(.ColumnHeaders.count).width = colWidth
            End If
        Next colName
        
        ' フィルタリングとデータの追加
        For Each tblRow In tbl.ListRows
            matchRow = True ' 初期状態ではすべての行を表示する前提
            
            ' フィルタ条件が設定されている場合、その条件に一致するかを確認
            If Len(filterCriteria) > 0 Then
                If tblRow.Range.Cells(1, tbl.ListColumns("name").index).value <> filterCriteria Then
                    matchRow = False ' 条件に一致しない場合は表示しない
                End If
            End If
            
            ' 条件に一致した場合のみリストに追加
            If matchRow Then
                Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, tbl.ListColumns(selectedColumns(1)).index).value)
                For i = 2 To selectedColumns.count
                    itm.SubItems(i - 1) = tblRow.Range.Cells(1, tbl.ListColumns(selectedColumns(i)).index).value
                Next i
            End If
        Next tblRow
    End With
End Sub

' コンボボックスの変更時に再描画
Private Sub ComboBox1_Change()
    ' コンボボックスの変更をトリガーに再度データを描画
    Call UserForm_Initialize
End Sub
