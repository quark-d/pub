```vb

' 定数の定義
Const ADDRESS_DELIMITER As String = "#" ' 必要に応じて変更
Const FIRST_DELIMITER As String = "." ' 必要に応じて変更
Const SECOND_DELIMITER As String = ";" ' 必要に応じて変更
Const COLUMN_NAME1 As String = "ID" ' 実際の列名に置き換えてください
Const COLUMN_NAME2 As String = "item3" ' 実際の列名に置き換えてください
Const SHEET_NAME1 As String = "Sheet2" ' 実際のシート名に置き換えてください
Const TABLE_NAME1 As String = "tbl_sht2" ' 実際のテーブル名に置き換えてください



Sub ProcessTableData()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim oneRecordList As Collection
    Dim oneRecord As clsOneRecord
    Dim addressStruct As clsAddress
    Dim column1 As String
    Dim column2 As String
    Dim splitData() As String
    Dim addressData() As String
    Dim i As Integer
    Dim j As Integer
    Dim item As Variant
    Dim visibleRows As Range
    Dim cell As Range
    Dim colIndex1 As Long
    Dim colIndex2 As Long
    
    ' 初期化
    Set ws = ThisWorkbook.Sheets(SHEET_NAME1)
    Set tbl = ws.ListObjects(TABLE_NAME1)
    Set oneRecordList = New Collection
    
    ' 列名から列インデックスを取得
    colIndex1 = tbl.ListColumns(COLUMN_NAME1).Index
    colIndex2 = tbl.ListColumns(COLUMN_NAME2).Index
    
    ' 可視セルの範囲を取得（フィルタ適用を考慮）
    On Error Resume Next
    Set visibleRows = tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' テーブルの可視セルを処理
    If Not visibleRows Is Nothing Then
        For Each cell In visibleRows.Columns(colIndex1).Cells
            column1 = cell.Value
            column2 = cell.Offset(0, colIndex2 - colIndex1).Value
            
            ' column2をADDRESS_DELIMITERで分割
            addressData = Split(column2, ADDRESS_DELIMITER)
            
            ' OneRecordを作成
            Set oneRecord = New clsOneRecord
            oneRecord.column1 = column1
            
            ' 各Addressを処理
            For i = LBound(addressData) To UBound(addressData)
                Set addressStruct = New clsAddress
                splitData = Split(addressData(i), FIRST_DELIMITER)
                
                ' address構造体を作成
                addressStruct.sheetName = splitData(0)
                
                ' アイテムリストを作成
                For j = 1 To UBound(splitData)
                    Dim itemParts() As String
                    itemParts = Split(splitData(j), SECOND_DELIMITER)
                    For Each item In itemParts
                        addressStruct.items.Add item
                    Next item
                Next j
                
                ' AddressをOneRecordに追加
                oneRecord.column2.Add addressStruct
            Next i
            
            ' OneRecordをリストに追加
            oneRecordList.Add oneRecord
        Next cell
    End If
    
    ' 任意のcolumn1を指定してaddressを取得
    Dim searchColumn1 As String
    searchColumn1 = "AAA-0003" ' 実際の値に置き換えてください
    
    Dim rec As clsOneRecord
    Dim existingShapes As Collection
    Set existingShapes = New Collection
    
    For Each rec In oneRecordList
        If rec.column1 = searchColumn1 Then
            Dim addr As clsAddress
            
            For Each addr In rec.column2
                ' シート名を基にシートの存在確認と作成
                On Error Resume Next
                Set ws = ThisWorkbook.Sheets(addr.sheetName)
                On Error GoTo 0
                If ws Is Nothing Then
                    Set ws = ThisWorkbook.Sheets.Add
                    ws.Name = addr.sheetName
                End If
                
                ' オートシェイプの名前作成
                Dim shapeName As String
                shapeName = addr.sheetName
                For Each item In addr.items
                    shapeName = shapeName & item
                    Debug.Print shapeName
                Next item
                
                ' オートシェイプが存在するかチェック
                Dim shapeExists As Boolean
                shapeExists = False
                Dim shp As Shape
                For Each shp In ws.Shapes
                    If shp.Name = shapeName Then
                        shapeExists = True
                        existingShapes.Add shp
                        shp.Visible = msoTrue
                        Exit For
                    End If
                Next shp
                
                ' オートシェイプを描画
                If Not shapeExists Then
                    Dim autoShape As Shape
                    Set autoShape = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 70, 25) ' 位置とサイズは適宜調整
                    autoShape.Name = shapeName
                    autoShape.TextFrame.Characters.Text = searchColumn1
                    existingShapes.Add autoShape
                End If
            Next addr
            
            Exit For
        End If
    Next rec
    
    ' 常に表示しておきたいオートシェイプのリストを定義
    Dim listNamesOfShapesToAlwaysDisplay As Collection
    Set listNamesOfShapesToAlwaysDisplay = New Collection
    listNamesOfShapesToAlwaysDisplay.Add "ShapeName1" ' 実際のシェイプ名に置き換えてください
    listNamesOfShapesToAlwaysDisplay.Add "ShapeName2" ' 実際のシェイプ名に置き換えてください
    ' 必要に応じて追加
    
    ' すべてのシートに対して存在しないオートシェイプを非表示にする
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        For Each shp In sheet.Shapes
            Dim found As Boolean
            found = False
            For Each item In existingShapes
                Debug.Print shp.Name & "<>" & item.Name
                If shp.Name = item.Name Then
                    found = True
                    Exit For
                End If
            Next item
            
            ' 常に表示しておきたいオートシェイプの名前をチェック
            If Not found Then
                For Each item In listNamesOfShapesToAlwaysDisplay
                    If shp.Name = item Then
                        found = True
                        Exit For
                    End If
                Next item
            End If
            
            If Not found Then
                shp.Visible = msoFalse
            End If
        Next shp
    Next sheet
End Sub
```
```vb
Public searchColumn1 As String

Sub ShowSelectionForm()
    searchColumn1 = ""
    frmSelection.Show
    If searchColumn1 <> "" Then
        ProcessTableData
    Else
        MsgBox "No selection made.", vbExclamation
    End If
End Sub
```
```vb
' clsAddress クラスモジュール
Public sheetName As String
Public items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub
```
```vb
' clsOneRecord クラスモジュール
Public column1 As String
Public column2 As Collection ' Collection of clsAddress

Private Sub Class_Initialize()
    Set column2 = New Collection
End Sub
```
```vb
' frmSelection
Option Explicit
Const ADDRESS_DELIMITER As String = "#" ' 必要に応じて変更
Const FIRST_DELIMITER As String = "." ' 必要に応じて変更
Const SECOND_DELIMITER As String = ";" ' 必要に応じて変更
Const COLUMN_NAME1 As String = "ID" ' 実際の列名に置き換えてください
Const COLUMN_NAME2 As String = "item3" ' 実際の列名に置き換えてください
Const SHEET_NAME1 As String = "Sheet2" ' 実際のシート名に置き換えてください
Const TABLE_NAME1 As String = "tbl_sht2" ' 実際のテーブル名に置き換えてください

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colIndex1 As Long
    Dim cell As Range
    Dim topPosition As Single
    Dim frameHeight As Single

    ' 初期化
    Set ws = ThisWorkbook.Sheets(SHEET_NAME1)
    Set tbl = ws.ListObjects(TABLE_NAME1)
    colIndex1 = tbl.ListColumns(COLUMN_NAME1).Index
    
    ' ラジオボタンの配置位置
    topPosition = 10
    frameHeight = 0

    ' COLUMN_NAME1の値をラジオボタンとして追加
    For Each cell In tbl.ListColumns(colIndex1).DataBodyRange.Cells
        If cell.Value <> "" Then
            Dim optButton As MSForms.OptionButton
            Set optButton = Me.Frame1.Controls.Add("Forms.OptionButton.1")
            optButton.Caption = cell.Value
            optButton.Left = 10
            optButton.Top = topPosition
            optButton.Width = 200
            topPosition = topPosition + 20
            frameHeight = frameHeight + 20
        End If
    Next cell
    
    ' フレームのサイズを調整
    Me.Frame1.Height = frameHeight + 20 ' 余白を追加
    Me.Height = Me.Frame1.Top + Me.Frame1.Height + 40 ' フォームの高さを調整
End Sub

Private Sub CommandButton1_Click()
    Dim ctrl As Control
    For Each ctrl In Me.Frame1.Controls
        If TypeName(ctrl) = "OptionButton" And ctrl.Value = True Then
            searchColumn1 = ctrl.Caption
            Exit For
        End If
    Next ctrl
    Me.Hide
End Sub

'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim colIndex1 As Long
'    Dim cell As Range
'    Dim topPosition As Single
'
'    ' 初期化
'    Set ws = ThisWorkbook.Sheets(SHEET_NAME1)
'    Set tbl = ws.ListObjects(TABLE_NAME1)
'    colIndex1 = tbl.ListColumns(COLUMN_NAME1).Index
'
'    ' ラジオボタンの配置位置
'    topPosition = 20
'
'    ' COLUMN_NAME1の値をラジオボタンとして追加
'    For Each cell In tbl.ListColumns(colIndex1).DataBodyRange.Cells
'        If cell.Value <> "" Then
'            Dim optButton As OptionButton
'            Set optButton = Me.Frame1.Controls.Add("Forms.OptionButton.1", "name", True)
'            optButton.Caption = cell.Value
'            optButton.Left = 10
'            optButton.Top = topPosition
'            optButton.Width = 200
'            topPosition = topPosition + 20
'        End If
'    Next cell
'End Sub
'
'Private Sub CommandButton1_Click()
'    Dim ctrl As Control
'    For Each ctrl In Me.Frame1.Controls
'        If TypeName(ctrl) = "OptionButton" And ctrl.Value = True Then
'            searchColumn1 = ctrl.Caption
'            Exit For
'        End If
'    Next ctrl
'    Me.Hide
'End Sub
'

```



