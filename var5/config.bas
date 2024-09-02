Option Explicit
' Class

' Public変数としてマーカーとラベルの情報を保持
Public SHAPE_POSITION_TYPE As MsoAutoShapeType
Public SHAPE_POSITION_WIDTH As Single
Public SHAPE_POSITION_HEIGHT As Single
Public SHAPE_POSITION_BK_COLOR_DEFAULT As Long
Public SHAPE_POSITION_BK_COLOR As Long
Public SHAPE_POSITION_BK_COLOR_ADD As Long
Public SHAPE_POSITION_BK_COLOR_DEL As Long
Public SHAPE_POSITION_FONT_COLOR As Long

Public SHAPE_LABEL_TYPE As MsoAutoShapeType
Public SHAPE_LABEL_WIDTH As Single
Public SHAPE_LABEL_HEIGHT As Single
Public SHAPE_LABEL_BK_COLOR As Long
Public SHAPE_LABEL_BK_COLOR_ADD As Long
Public SHAPE_LABEL_BK_COLOR_DEL As Long
Public SHAPE_LABEL_FONT_COLOR As Long

Private configDict As Scripting.Dictionary
Private CONFIG_SHEET_NAME As String
Private CONFIG_TABLE_NAME As String

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
    Dim parameterName As String
    Dim valueType As String
    Dim value As Variant
    
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
        ws.Range("A1:E1").Value = Array("Parameter", "Value", "RGB_R", "RGB_G", "RGB_B")
        ' テーブルを作成
        Set configTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E1"), , xlYes)
        configTable.Name = CONFIG_TABLE_NAME
        
        ' 初期データをセット（省略しますが、既存コードと同じ）
    End If
    
    ' テーブルから設定データを動的に読み込む
    Dim row As ListRow
    For Each row In configTable.ListRows
        parameterName = row.Range(1, 1).Value ' "Parameter" 列の値を取得
        valueType = "Value" ' デフォルトの型は "Value" とする
        
        ' RGBの場合、RGB値を取得
        If Not IsEmpty(row.Range(1, 3).Value) Then
            value = RGB(row.Range(1, 3).Value, row.Range(1, 4).Value, row.Range(1, 5).Value)
            valueType = "RGB"
        Else
            value = row.Range(1, 2).Value ' "Value" 列の値を取得
        End If
        
        ' パラメータ名に基づいて対応する変数に値をセット
        Select Case parameterName
            Case "SHAPE_POSITION_TYPE"
                SHAPE_POSITION_TYPE = GetShapeType(value)
            Case "SHAPE_POSITION_WIDTH"
                SHAPE_POSITION_WIDTH = value
            Case "SHAPE_POSITION_HEIGHT"
                SHAPE_POSITION_HEIGHT = value
            Case "SHAPE_POSITION_BK_COLOR_DEFAULT"
                SHAPE_POSITION_BK_COLOR_DEFAULT = value
            Case "SHAPE_POSITION_BK_COLOR"
                SHAPE_POSITION_BK_COLOR = value
            Case "SHAPE_POSITION_BK_COLOR_ADD"
                SHAPE_POSITION_BK_COLOR_ADD = value
            Case "SHAPE_POSITION_BK_COLOR_DEL"
                SHAPE_POSITION_BK_COLOR_DEL = value
            Case "SHAPE_POSITION_FONT_COLOR"
                SHAPE_POSITION_FONT_COLOR = value
            Case "SHAPE_LABEL_TYPE"
                SHAPE_LABEL_TYPE = GetShapeType(value)
            Case "SHAPE_LABEL_WIDTH"
                SHAPE_LABEL_WIDTH = value
            Case "SHAPE_LABEL_HEIGHT"
                SHAPE_LABEL_HEIGHT = value
            Case "SHAPE_LABEL_BK_COLOR"
                SHAPE_LABEL_BK_COLOR = value
            Case "SHAPE_LABEL_BK_COLOR_ADD"
                SHAPE_LABEL_BK_COLOR_ADD = value
            Case "SHAPE_LABEL_BK_COLOR_DEL"
                SHAPE_LABEL_BK_COLOR_DEL = value
            Case "SHAPE_LABEL_FONT_COLOR"
                SHAPE_LABEL_FONT_COLOR = value
            ' 追加のパラメータがある場合はここに追加
        End Select
    Next row
End Sub

Private Function GetShapeType(shapeTypeName As String) As MsoAutoShapeType
    Select Case shapeTypeName
        Case "msoShapeRectangle"
            GetShapeType = msoShapeRectangle
        Case "msoShapeOval"
            GetShapeType = msoShapeOval
        Case "msoShapeRoundedRectangle"
            GetShapeType = msoShapeRoundedRectangle
        Case "msoTextBox"
            GetShapeType = msoTextBox
        Case Else
            GetShapeType = msoShapeRectangle ' デフォルト値
    End Select
End Function

Public Function GetConfigShapePosition() As Dictionary
    Dim shapeConfig As New Dictionary
    shapeConfig.Add "SHAPE_POSITION_TYPE", SHAPE_POSITION_TYPE
    shapeConfig.Add "SHAPE_POSITION_WIDTH", SHAPE_POSITION_WIDTH
    shapeConfig.Add "SHAPE_POSITION_HEIGHT", SHAPE_POSITION_HEIGHT
    shapeConfig.Add "SHAPE_POSITION_BK_COLOR_DEFAULT", SHAPE_POSITION_BK_COLOR_DEFAULT
    shapeConfig.Add "SHAPE_POSITION_BK_COLOR", SHAPE_POSITION_BK_COLOR
    shapeConfig.Add "SHAPE_POSITION_BK_COLOR_ADD", SHAPE_POSITION_BK_COLOR_ADD
    shapeConfig.Add "SHAPE_POSITION_BK_COLOR_DEL", SHAPE_POSITION_BK_COLOR_DEL
    shapeConfig.Add "SHAPE_POSITION_FONT_COLOR", SHAPE_POSITION_FONT_COLOR
    shapeConfig.Add "SHAPE_LABEL_TYPE", SHAPE_LABEL_TYPE
    shapeConfig.Add "SHAPE_LABEL_WIDTH", SHAPE_LABEL_WIDTH
    shapeConfig.Add "SHAPE_LABEL_HEIGHT", SHAPE_LABEL_HEIGHT
    shapeConfig.Add "SHAPE_LABEL_BK_COLOR", SHAPE_LABEL_BK_COLOR
    shapeConfig.Add "SHAPE_LABEL_BK_COLOR_ADD", SHAPE_LABEL_BK_COLOR_ADD
    shapeConfig.Add "SHAPE_LABEL_BK_COLOR_DEL", SHAPE_LABEL_BK_COLOR_DEL
    shapeConfig.Add "SHAPE_LABEL_FONT_COLOR", SHAPE_LABEL_FONT_COLOR
    Set GetConfigShapePosition = shapeConfig
End Function
