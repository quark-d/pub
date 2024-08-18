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

Sub EntryPointMulti(ByVal selectedIDs As Collection)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim individualPosition As Variant
    Dim positionList As New Collection
    Dim selectedID As Variant
    
    ' configPositionの初期化
    Call InitializeconfigPosition
    
    Dim idAddrDict As Scripting.Dictionary
    Set idAddrDict = GetIdAddrDictionary() ' IDを取得する関数を呼び出す
    
    ' 複数のselectedIDを処理
    For Each selectedID In selectedIDs
        If Not idAddrDict.Exists(selectedID) Then
            MsgBox "指定されたID " & selectedID & " は辞書に存在しません。", vbExclamation
            Exit Sub
        End If
        
        Dim addr As String
        addr = idAddrDict(selectedID)
        
        ' 位置リストを取得し、positionListに追加
        Dim tempPositionList As Collection
        Set tempPositionList = GetLocationListSingle(addr)
        
        For Each individualPosition In tempPositionList
            positionList.Add individualPosition
        Next individualPosition
    Next selectedID
    
    ' 位置をセットアップ
    For Each individualPosition In positionList
        configPosition.SetUpPosition CStr(individualPosition)
    Next individualPosition
    
    ' DrowPositions 関数を呼び出す
    Call DrowShapes(configShape, configPosition, positionList)
    
    ' シートとテーブルの設定
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    Call AppendTextInLabel(tbl, configShape, selectedIDs)

    ' ラベルのサイズを調整
    Dim labelShape As shape
    For Each labelShape In GetAutoShapeListLabel()
        Call AdjustShapeLabelByText(labelShape)
    Next labelShape
End Sub



Sub EntryPointOne(ByVal selectedID As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim individualPosition As Variant
    Dim positionList As Collection
    
    Call InitializeconfigPosition
    
    
    Dim idAddrDict As Scripting.Dictionary
    Set idAddrDict = GetIdAddrDictionary() ' IDを取得する関数を呼び出す
    If Not idAddrDict.Exists(selectedID) Then
        MsgBox "指定されたselectedIDは辞書に存在しません。", vbExclamation
        Exit Sub
    End If
    
    Dim addr As String
    addr = idAddrDict(selectedID)
    ' 位置リストを取得
    Set positionList = GetLocationListSingle(addr)
    
    For Each individualPosition In positionList
        configPosition.SetUpPosition CStr(individualPosition)
    Next individualPosition
    
    ' DrowPositions 関数を呼び出す
    Call DrowShapes(configShape, configPosition, positionList)
    
    ' シートとテーブルの設定
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    Dim selectedIDs As New Collection ' Collectionを作成
    selectedIDs.Add selectedID
    Call AppendTextInLabel(tbl, configShape, selectedIDs)

    ' ラベルのサイズを調整
    Dim labelShape As shape
    For Each labelShape In GetAutoShapeListLabel()
        Call AdjustShapeLabelByText(labelShape)
    Next labelShape
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
    
    Dim row As ListRow
    Dim idCollection As New Collection ' IDを格納するCollectionを作成
    ' テーブルからID列の値を取得し、Collectionに追加
    For Each row In tbl.ListRows
        idCollection.Add row.Range(1).value ' IDが含まれている列を指定（ここでは1列目）
    Next row
    
    Call AppendTextInLabel(tbl, configShape, idCollection)

    ' ラベルのサイズを調整
    Dim labelShape As shape
    For Each labelShape In GetAutoShapeListLabel()
        Call AdjustShapeLabelByText(labelShape)
    Next labelShape
End Sub


Function GetLocationListSingle(addr As String) As Collection
    Dim positionList As New Collection
    Dim positions() As String
    Dim individualPosition As String
    Dim i As Integer
    Dim alreadyExists As Boolean
    Dim item As Variant
    
    ' 位置情報を ":" で分割
    positions = Split(addr, ":")
    
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
    
    ' コレクションを返す
    Set GetLocationListSingle = positionList
End Function

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

Function GetIdAddrDictionary() As Scripting.Dictionary
    Dim idAddrDict As New Scripting.Dictionary
    Dim row As ListRow
    Dim id As String
    Dim addr As String
    
    ' シートとテーブルの設定
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")

    ' テーブルからIDとaddrのペアを取得して辞書に保存
    For Each row In tbl.ListRows
        id = row.Range(tbl.ListColumns("ID").Index).value
        addr = row.Range(tbl.ListColumns("addr").Index).value
        idAddrDict.Add id, addr
    Next row
    
    Set GetIdAddrDictionary = idAddrDict
End Function
