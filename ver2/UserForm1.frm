VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim header As String
'    Dim i As Long
'    Dim dataRow As ListRow
'
'    ' シートとテーブルを取得
'    Set ws = ThisWorkbook.Sheets("map") ' シート名を指定
'    Set tbl = ws.ListObjects("tbl_maps")
'
'    ' ヘッダをListBoxに追加
'    header = ""
'    For i = 1 To tbl.HeaderRowRange.Columns.Count
'        If i > 1 Then header = header & vbTab ' タブで区切る
'        header = header & tbl.HeaderRowRange.Cells(1, i).value
'    Next i
'    Me.ListBox1.AddItem header
'
'    ' テーブルのデータをListBoxに追加
'    For Each dataRow In tbl.ListRows
'        rowData = ""
'        For i = 1 To dataRow.Range.Columns.Count
'            If i > 1 Then rowData = rowData & vbTab ' タブで区切る
'            rowData = rowData & dataRow.Range.Cells(1, i).value
'        Next i
'        Me.ListBox1.AddItem rowData
'    Next dataRow
'End Sub


'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim dataRow As ListRow
'    Dim i As Long
'    Dim lbl As MSForms.label
'    Dim leftPos As Single
'    Dim columnWidth As Single
'    Dim currentWidth As Single
'
'    ' シートとテーブルを取得
'    Set ws = ThisWorkbook.Sheets("map") ' シート名を指定
'    Set tbl = ws.ListObjects("tbl_maps")
'
'    ' 列幅を動的に計算 (ListBoxの幅を列数で割る)
'    columnWidth = Me.ListBox1.width / tbl.HeaderRowRange.Columns.Count
'
'    ' 初期の左位置を設定
'    leftPos = Me.ListBox1.left
'
'    ' ヘッダを表すラベルを動的に生成
'    For i = 1 To tbl.HeaderRowRange.Columns.Count
'        Set lbl = Me.Controls.Add("Forms.Label.1")
'
'        With lbl
'            .Caption = tbl.HeaderRowRange.Cells(1, i).value
'            .width = columnWidth ' 計算した列幅を設定
'            ' 列の中央にラベルを配置
'            .left = leftPos + (columnWidth - .width) / 2
'            .top = Me.ListBox1.top - 15
'            .TextAlign = fmTextAlignCenter
'        End With
'
'        ' 次のラベルの左位置を調整
'        leftPos = leftPos + columnWidth
'    Next i
'
'    ' ListBoxの設定
'    With Me.ListBox1
'        .ColumnCount = tbl.HeaderRowRange.Columns.Count ' テーブルの列数を設定
'        .BoundColumn = 1
'        .ColumnHeads = False
'    End With
'
'    ' データをListBoxに追加する
'    For Each dataRow In tbl.ListRows
'        Dim rowData(1 To 10) As String
'        For i = 1 To tbl.HeaderRowRange.Columns.Count
'            rowData(i) = dataRow.Range.Cells(1, i).value
'        Next i
'        Me.ListBox1.AddItem Join(rowData, vbTab)
'    Next dataRow
'End Sub
''
'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRow As ListRow
'    Dim dataArray() As String
'    Dim i As Long
'
'    ' ワークシートとテーブルを指定
'    Set ws = ThisWorkbook.Sheets("tbl_maps")
'    Set tbl = ws.ListObjects("tbl_maps_1")
'
'    ' ヘッダを追加
'    With Me.ListBox1
'        .ColumnCount = tbl.ListColumns.Count
'        .AddItem "ID" & vbTab & "name" & vbTab & "addr"
'
'        ' データを追加
'        For Each tblRow In tbl.ListRows
'            ReDim dataArray(1 To tbl.ListColumns.Count)
'            For i = 1 To tbl.ListColumns.Count
'                dataArray(i) = tblRow.Range(1, i).value
'            Next i
'            .AddItem Join(dataArray, vbTab)
'        Next tblRow
'    End With
'End Sub

'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRow As ListRow
'    Dim itm As ListItem
'    Dim i As Integer
'    Dim colWidth As Integer
'
'    ' ワークシートとテーブルを指定
'    Set ws = ThisWorkbook.Sheets("maps")
'    Set tbl = ws.ListObjects("tbl_maps_1")
'
'    ' ListViewの設定
'    With Me.ListView1
'        .View = lvwReport
'        .Gridlines = True
'        .FullRowSelect = True
'        .ColumnHeaders.Clear
'        .ListItems.Clear
'
'        ' テーブルのヘッダを追加
'        For i = 1 To tbl.ListColumns.Count
'            .ColumnHeaders.Add , , tbl.ListColumns(i).Name, 100 ' 初期幅を設定
'        Next i
'
'        ' テーブルのデータを追加
'        For Each tblRow In tbl.ListRows
'            Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, 1).value) ' 最初の列のデータ
'            For i = 2 To tbl.ListColumns.Count
'                itm.SubItems(i - 1) = tblRow.Range.Cells(1, i).value
'            Next i
'        Next tblRow
'
'        ' 列幅を手動で調整
'        For i = 1 To .ColumnHeaders.Count
'            colWidth = .ColumnHeaders(i).width
'            ' データとヘッダの長さに基づいて列幅を設定
'            For Each itm In .ListItems
'                colWidth = Application.Max(colWidth, Me.TextWidth(itm.Text) + 10)
'            Next itm
'            .ColumnHeaders(i).width = colWidth
'        Next i
'    End With
'End Sub


'Dim savedSelections As Collection
'
'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRow As ListRow
'    Dim itm As ListItem
'    Dim i As Integer
'
'    ' ワークシートとテーブルを指定
'    Set ws = ThisWorkbook.Sheets("tbl_maps")
'    Set tbl = ws.ListObjects("tbl_maps_1")
'
'    ' 選択状態を保存するためのコレクションを初期化
'    Set savedSelections = New Collection
'
'    ' ListViewの設定
'    With Me.ListView1
'        .View = lvwReport
'        .Gridlines = True
'        .FullRowSelect = True
'        .MultiSelect = True ' 複数選択を有効にする
'        .ColumnHeaders.Clear
'        .ListItems.Clear
'
'        ' テーブルのヘッダを追加
'        For i = 1 To tbl.ListColumns.Count
'            .ColumnHeaders.Add , , tbl.ListColumns(i).Name, 100 ' 固定幅を設定
'        Next i
'
'        ' テーブルのデータを追加
'        For Each tblRow In tbl.ListRows
'            Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, 1).value) ' 最初の列のデータ
'            For i = 2 To tbl.ListColumns.Count
'                itm.SubItems(i - 1) = tblRow.Range.Cells(1, i).value
'            Next i
'        Next tblRow
'    End With
'End Sub
'
'Private Sub UserForm_Deactivate()
'    ' 選択されたアイテムを保存
'    Dim itm As ListItem
'    Set savedSelections = New Collection
'    For Each itm In Me.ListView1.selectedItems
'        savedSelections.Add itm.Text
'    Next itm
'End Sub
'
'Private Sub UserForm_Activate()
'    ' 保存された選択を復元
'    Dim itm As ListItem
'    Dim selItem As Variant
'
'    If savedSelections.Count > 0 Then
'        For Each selItem In savedSelections
'            For Each itm In Me.ListView1.ListItems
'                If itm.Text = selItem Then
'                    itm.Selected = True
'                    Exit For
'                End If
'            Next itm
'        Next selItem
'    End If
'End Sub

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

    ' 表示する列名をCollectionに追加
    selectedColumns.Add "ID"
    selectedColumns.Add "name"
    selectedColumns.Add "addr"
    selectedColumns.Add "addr_add"
    ' 必要な列名を追加

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
                .ColumnHeaders(.ColumnHeaders.Count).width = colWidth
            End If
        Next colName
        
        ' テーブルのデータを追加
        For Each tblRow In tbl.ListRows
            Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, tbl.ListColumns(selectedColumns(1)).index).value)
            For i = 2 To selectedColumns.Count
                itm.SubItems(i - 1) = tblRow.Range.Cells(1, tbl.ListColumns(selectedColumns(i)).index).value
            Next i
        Next tblRow
    End With
End Sub

Private Sub UserForm_Terminate()
    Dim wsConfig As Worksheet
    Dim i As Integer
    
    ' 設定シートを取得
    Set wsConfig = ThisWorkbook.Sheets("ConfigListView")
    
    ' 列幅を保存
    With Me.ListView1
        For i = 1 To .ColumnHeaders.Count
            wsConfig.Cells(1, i).value = .ColumnHeaders(i).width
        Next i
    End With
End Sub



Private Sub CommandButton1_Click()
    Dim itm As ListItem
    Dim checkedItems As Collection
    Set checkedItems = New Collection

    ' チェックされているアイテムのテキストをコレクションに追加
    For Each itm In Me.ListView1.ListItems
        If itm.Checked Then
            checkedItems.Add itm.Text
        End If
    Next itm

    ' コレクションの内容をDebug.Printで出力
    Dim item As Variant
    For Each item In checkedItems
        Debug.Print item
    Next item
End Sub

