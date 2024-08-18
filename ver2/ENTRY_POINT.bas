Attribute VB_Name = "ENTRY_POINT"
Option Explicit


Dim configShape As ConfigShapes
Dim configPosition As ConfigPositions

Sub InitializeconfigPosition()
    Set configShape = New ConfigShapes
    configShape.LoadConfigData "ConfigShapes", "tbl_config_shapes"

    Set configPosition = New ConfigPositions
    configPosition.LoadConfigData "ConfigPositions", "tbl_config_positions"
   
End Sub

Sub EntryPoint()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim individualPosition As Variant
    Dim positionList As Collection
    
    Call InitializeconfigPosition
    
    ' シートとテーブルの設定
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    
    ' 位置リストを取得
    Set positionList = GetLocationList(tbl)
    
    For Each individualPosition In positionList
        configPosition.SetUpPosition CStr(individualPosition)
    Next individualPosition
    
    ' DrowPositions 関数を呼び出す
    Call DrowShapes(configShape, configPosition, positionList)
    
    Call AppendTextInLabel(tbl, configShape)

    ' ラベルのサイズを調整
    Dim labelShape As shape
    For Each labelShape In GetAutoShapeListLabel()
        Call AdjustShapeLabelByText(labelShape)
    Next labelShape
    
    
    
End Sub


Function GetLocationList(tbl As ListObject) As Collection
    Dim positionList As New Collection
    Dim row As ListRow
    Dim positions() As String
    Dim individualPosition As String
    Dim i As Integer
    Dim alreadyExists As Boolean
    Dim item As Variant
    
    ' テーブルの各行をループ
    For Each row In tbl.ListRows
        ' addr 列から位置情報を取得
        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, ":")
        
        ' それぞれの位置をループしてコレクションに追加
        For i = LBound(positions) To UBound(positions)
            individualPosition = positions(i)
            alreadyExists = False
            
            ' コレクション内に既に存在するかチェック
            For Each item In positionList
                If item = individualPosition Then
                    alreadyExists = True
                    Exit For
                End If
            Next item
            
            ' 既に存在しなければコレクションに追加
            If Not alreadyExists Then
                positionList.Add individualPosition
            End If
        Next i
    Next row
    
    ' コレクションを返す
    Set GetLocationList = positionList
End Function

