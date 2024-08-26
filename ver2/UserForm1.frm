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

