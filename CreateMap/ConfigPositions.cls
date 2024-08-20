VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dim positionsConfig As ConfigPositions
'
'Sub InitializePositionsConfig()
'    Set positionsConfig = New ConfigPositions
'    positionsConfig.LoadConfigData "configPosition", "tbl_config_positions"
'End Sub
'
'Sub BeforeWorkbookClose(Cancel As Boolean)
'    InitializePositionsConfig
'    If positionsConfig.CheckForPositionChanges Then
'        MsgBox "位置が保存されました。"
'    End If
'End Sub

Private configDict As Scripting.Dictionary
Private CONFIG_SHEET_NAME As String
Private CONFIG_TABLE_NAME As String

Private Const DELIMITER_ADDR As String = "/"

Private Sub Class_Initialize()
    Set configDict = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    ' 必要に応じてクリーンアップ処理を実装
End Sub

Public Sub LoadConfigData(sheetName As String, tableName As String)
    ' シート名とテーブル名を private 変数に保存
    CONFIG_SHEET_NAME = sheetName
    CONFIG_TABLE_NAME = tableName

    Dim ws As Worksheet
    Dim configTable As ListObject
    
    ' シートが存在しない場合は作成
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = CONFIG_SHEET_NAME
    End If
    
    ' テーブルが存在しない場合は作成
    On Error Resume Next
    Set configTable = ws.ListObjects(CONFIG_TABLE_NAME)
    On Error GoTo 0
    If configTable Is Nothing Then
        ' ヘッダー行をセット
        ws.Range("A1:F1").value = Array("ShapeID", "X", "Y", "RGB_R", "RGB_G", "RGB_B")
        ' テーブルを作成
        Set configTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F1"), , xlYes)
        configTable.Name = CONFIG_TABLE_NAME
    End If
    
    ' テーブルから設定データを読み込む
    Dim row As ListRow
    For Each row In configTable.ListRows
        Dim shapeID As String
        shapeID = row.Range(configTable.ListColumns("ShapeID").Index).value
        If Len(shapeID) > 0 Then
            configDict.Add shapeID, Array( _
                row.Range(configTable.ListColumns("X").Index).value, _
                row.Range(configTable.ListColumns("Y").Index).value, _
                RGB( _
                    row.Range(configTable.ListColumns("RGB_R").Index).value, _
                    row.Range(configTable.ListColumns("RGB_G").Index).value, _
                    row.Range(configTable.ListColumns("RGB_B").Index).value _
                ) _
            )
        End If
    Next row
End Sub

Public Sub SaveConfigData()
    Dim ws As Worksheet
    Dim configTable As ListObject

    ' シートとテーブルを取得
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    Set configTable = ws.ListObjects(CONFIG_TABLE_NAME)
    
    ' テーブルをクリア
    configTable.DataBodyRange.ClearContents
    
    ' 辞書からデータをテーブルに書き戻す
    Dim shapeID As Variant
    For Each shapeID In configDict.Keys
        With configTable.ListRows.Add
            .Range(configTable.ListColumns("ShapeID").Index).value = shapeID
            .Range(configTable.ListColumns("X").Index).value = configDict(shapeID)(0)
            .Range(configTable.ListColumns("Y").Index).value = configDict(shapeID)(1)
            .Range(configTable.ListColumns("RGB_R").Index).value = Red(configDict(shapeID)(2))
            .Range(configTable.ListColumns("RGB_G").Index).value = Green(configDict(shapeID)(2))
            .Range(configTable.ListColumns("RGB_B").Index).value = Blue(configDict(shapeID)(2))
        End With
    Next shapeID
End Sub

Public Sub AddOrUpdatePosition(shapeID As String, posX As Single, posY As Single, bkColor As Long)
    ' 位置情報を追加または更新する
    If configDict.Exists(shapeID) Then
        configDict(shapeID) = Array(posX, posY, bkColor)
    Else
        configDict.Add shapeID, Array(posX, posY, bkColor)
    End If
End Sub

Public Sub SetUpPosition(positionString As String)
    Dim ws As Worksheet
    Dim posWs As Worksheet
    Dim tbl As ListObject
    Dim shp As shape
    Dim shapeName As String
    Dim sheetName As String
    Dim position As String
    Dim colorRGB As Long
    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer

    ' シートと位置情報を分割して取得
    sheetName = Split(positionString, DELIMITER_ADDR)(0)
    position = Split(positionString, DELIMITER_ADDR)(1)
    shapeName = "pos_" & sheetName & "_" & position

    ' 位置情報シート
    Set posWs = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    ' テーブル
    Set tbl = posWs.ListObjects(CONFIG_TABLE_NAME)
    
    ' 該当シートの確認または作成
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    End If
    On Error GoTo 0

    ' 規定のRGB色を取得
    On Error Resume Next
    ' ダミーシェイプを作成して規定の色を取得
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
    colorRGB = shp.Fill.ForeColor.RGB
    redValue = colorRGB And 255
    greenValue = (colorRGB \ 256) And 255
    blueValue = (colorRGB \ 65536) And 255
    shp.Delete ' ダミーシェイプを削除
    On Error GoTo 0

    ' 既存のレコードを検索
    Dim foundRow As ListRow
    For Each foundRow In tbl.ListRows
        If foundRow.Range(1).value = shapeName Then Exit Sub
    Next foundRow

    ' 新規レコードを追加
    With tbl.ListRows.Add
        .Range(1).value = shapeName
        .Range(2).value = 0 ' 初期X座標
        .Range(3).value = 0 ' 初期Y座標
        .Range(4).value = redValue
        .Range(5).value = greenValue
        .Range(6).value = blueValue
    End With
End Sub

Public Function GetTblObject() As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    Set GetTblObject = ws.ListObjects(CONFIG_TABLE_NAME)
End Function

Public Sub AllocateShapeAttribute(shapeID As String, ByRef xPos As Double, ByRef yPos As Double, ByRef redValue As Integer, ByRef greenValue As Integer, ByRef blueValue As Integer)
    Dim foundRow As ListRow
    
    ' ConfigPositions テーブルから ShapeID を検索
    For Each foundRow In Me.GetTblObject().ListRows
        If foundRow.Range(1).value = shapeID Then
            xPos = foundRow.Range(2).value
            yPos = foundRow.Range(3).value
            redValue = foundRow.Range(4).value
            greenValue = foundRow.Range(5).value
            blueValue = foundRow.Range(6).value
            Exit Sub
        End If
    Next foundRow
End Sub




Private Function ExistsInCollection(value As Variant, col As Collection) As Boolean
    ' Collectionに特定の値が存在するかを確認するヘルパー関数
    Dim item As Variant
    For Each item In col
        If item = value Then
            ExistsInCollection = True
            Exit Function
        End If
    Next item
    ExistsInCollection = False
End Function


