Attribute VB_Name = "ENTRY_POINT"
Option Explicit

Private Const QUERY_SHEET As String = "map"
Private Const QUERY_TABLE As String = "tbl_maps"
Private Const QUERY_ITEM_ID As String = "ID"
Private Const QUERY_ITEM_ADDR As String = "addr"
Private Const DELIMITER_ADDRS As String = ":"

Private Const CONFIG_SHEET_SHAPES As String = "ConfigShapes"
Private Const CONFIG_TABLE_SHAPES As String = "tbl_config_shapes"
Private Const CONFIG_SHEET_POSITIONS As String = "ConfigPositions"
Private Const CONFIG_TABLE_POSITIONS As String = "tbl_config_positions"


Dim configShape As ConfigShapes
Dim configPosition As ConfigPositions

Sub InitializeconfigPosition()
    Set configShape = New ConfigShapes
    configShape.LoadConfigData CONFIG_SHEET_SHAPES, CONFIG_TABLE_SHAPES

    Set configPosition = New ConfigPositions
    configPosition.LoadConfigData CONFIG_SHEET_POSITIONS, CONFIG_TABLE_POSITIONS
   
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)
    
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
    positions = Split(addr, DELIMITER_ADDRS)
    
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
        positions = Split(row.Range(tbl.ListColumns(QUERY_ITEM_ADDR).Index).value, DELIMITER_ADDRS)
        
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)

    ' テーブルからIDとaddrのペアを取得して辞書に保存
    For Each row In tbl.ListRows
        id = row.Range(tbl.ListColumns(QUERY_ITEM_ID).Index).value
        addr = row.Range(tbl.ListColumns(QUERY_ITEM_ADDR).Index).value
        idAddrDict.Add id, addr
    Next row
    
    Set GetIdAddrDictionary = idAddrDict
End Function

' Public Function GetRoomsFromAddrDict(idAddrDict As Scripting.Dictionary) As Collection
Public Function GetSheetNamesFromAddrDict(idAddrDict As Scripting.Dictionary) As Collection
    Dim sheetNames As New Collection
    Dim addr As String
    Dim addrParts() As String
    Dim classroomParts() As String
    Dim sheetName As String
    Dim item As Variant
    Dim alreadyExists As Boolean
    Dim i As Long
    Dim existingSheet As Variant
    
    ' 辞書内の各アイテムを反復処理
    For Each item In idAddrDict.Items
        ' addr文字列を ":" で分割
        addrParts = Split(item, ":")
        
        ' 各部分を反復処理して、シート名を抽出
        For i = LBound(addrParts) To UBound(addrParts)
            classroomParts = Split(addrParts(i), "/")
            sheetName = classroomParts(0)
            
            ' コレクション内でシート名が既に存在するか確認
            alreadyExists = False
            For Each existingSheet In sheetNames
                If existingSheet = sheetName Then
                    alreadyExists = True
                    Exit For
                End If
            Next existingSheet
            
            ' シート名がユニークであればコレクションに追加
            If Not alreadyExists Then
                sheetNames.Add sheetName
            End If
        Next i
    Next item
    
    ' シート名のコレクションを返す
    Set GetSheetNamesFromAddrDict = sheetNames
End Function


Public Sub ResetAllSheetTabColors()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Tab.colorIndex = xlColorIndexNone
    Next ws
End Sub
Public Sub SetSheetTabColor(ws As Worksheet, Optional colorIndex As Long = 3)
    ' 指定されたシートのタブの色を colorIndex で設定
    ws.Tab.colorIndex = colorIndex
End Sub

