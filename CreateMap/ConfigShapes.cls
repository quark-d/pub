VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Scripting.Dictionaryを利用しているため
' VBAの参照設定で Microsoft Scripting Runtime を有効にする

'Dim shapeConfig As configShape
'
'Sub InitializeShapeConfig()
'    Set shapeConfig = New configShape
'    shapeConfig.LoadConfigData "configShape", "tbl_config_position"
'End Sub
'
'Sub ExampleUsage()
'    Call InitializeShapeConfig
'
'    MsgBox "Shape Type: " & shapeConfig.SHAPE_POSITION_TYPE & vbCrLf & _
'           "Width: " & shapeConfig.SHAPE_POSITION_WIDTH & vbCrLf & _
'           "Height: " & shapeConfig.SHAPE_POSITION_HEIGHT & vbCrLf & _
'           "Background Color: " & shapeConfig.SHAPE_POSITION_BK_COLOR
'End Sub

' Public変数としてマーカーとラベルの情報を保持
Public SHAPE_POSITION_TYPE As MsoAutoShapeType
Public SHAPE_POSITION_WIDTH As Single
Public SHAPE_POSITION_HEIGHT As Single
Public SHAPE_POSITION_BK_COLOR As Long

Public SHAPE_LABEL_TYPE As String
Public SHAPE_LABEL_WIDTH As Single
Public SHAPE_LABEL_HEIGHT As Single
Public SHAPE_LABEL_BK_COLOR As Long

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
        ws.Range("A1:E1").value = Array("Parameter", "Value", "RGB_R", "RGB_G", "RGB_B")
        ' テーブルを作成
        Set configTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E1"), , xlYes)
        configTable.Name = CONFIG_TABLE_NAME
        
        ' 初期データをセット
        configTable.ListRows.Add
        With configTable.ListRows(1)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_POSITION_TYPE"
            .Range(configTable.ListColumns("Value").Index).value = "msoShapeRectangle"
        End With
        configTable.ListRows.Add
        With configTable.ListRows(2)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_POSITION_WIDTH"
            .Range(configTable.ListColumns("Value").Index).value = 100
        End With
        configTable.ListRows.Add
        With configTable.ListRows(3)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_POSITION_HEIGHT"
            .Range(configTable.ListColumns("Value").Index).value = 50
        End With
        configTable.ListRows.Add
        With configTable.ListRows(4)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_POSITION_BK_COLOR"
            .Range(configTable.ListColumns("RGB_R").Index).value = 255
            .Range(configTable.ListColumns("RGB_G").Index).value = 0
            .Range(configTable.ListColumns("RGB_B").Index).value = 0
        End With
        
        ' ラベルの初期データをセット
        configTable.ListRows.Add
        With configTable.ListRows(5)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_LABEL_TYPE"
            .Range(configTable.ListColumns("Value").Index).value = "TextBox"
        End With
        configTable.ListRows.Add
        With configTable.ListRows(6)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_LABEL_WIDTH"
            .Range(configTable.ListColumns("Value").Index).value = 80
        End With
        configTable.ListRows.Add
        With configTable.ListRows(7)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_LABEL_HEIGHT"
            .Range(configTable.ListColumns("Value").Index).value = 30
        End With
        configTable.ListRows.Add
        With configTable.ListRows(8)
            .Range(configTable.ListColumns("Parameter").Index).value = "SHAPE_LABEL_BK_COLOR"
            .Range(configTable.ListColumns("RGB_R").Index).value = 200
            .Range(configTable.ListColumns("RGB_G").Index).value = 200
            .Range(configTable.ListColumns("RGB_B").Index).value = 200
        End With
    End If
    
    ' テーブルから設定データを読み込む
    SHAPE_POSITION_TYPE = GetShapeType(configTable.ListRows(1).Range(configTable.ListColumns("Value").Index).value)
    SHAPE_POSITION_WIDTH = configTable.ListRows(2).Range(configTable.ListColumns("Value").Index).value
    SHAPE_POSITION_HEIGHT = configTable.ListRows(3).Range(configTable.ListColumns("Value").Index).value
    SHAPE_POSITION_BK_COLOR = RGB(configTable.ListRows(4).Range(configTable.ListColumns("RGB_R").Index).value, _
                                  configTable.ListRows(4).Range(configTable.ListColumns("RGB_G").Index).value, _
                                  configTable.ListRows(4).Range(configTable.ListColumns("RGB_B").Index).value)
    
    SHAPE_LABEL_TYPE = GetShapeType(configTable.ListRows(5).Range(configTable.ListColumns("Value").Index).value)
    SHAPE_LABEL_WIDTH = configTable.ListRows(6).Range(configTable.ListColumns("Value").Index).value
    SHAPE_LABEL_HEIGHT = configTable.ListRows(7).Range(configTable.ListColumns("Value").Index).value
    SHAPE_LABEL_BK_COLOR = RGB(configTable.ListRows(8).Range(configTable.ListColumns("RGB_R").Index).value, _
                               configTable.ListRows(8).Range(configTable.ListColumns("RGB_G").Index).value, _
                               configTable.ListRows(8).Range(configTable.ListColumns("RGB_B").Index).value)
End Sub

Public Function GetConfigShapePosition() As Dictionary
    Dim shapeConfig As New Dictionary
    shapeConfig.Add "SHAPE_POSITION_TYPE", SHAPE_POSITION_TYPE
    shapeConfig.Add "SHAPE_POSITION_WIDTH", SHAPE_POSITION_WIDTH
    shapeConfig.Add "SHAPE_POSITION_HEIGHT", SHAPE_POSITION_HEIGHT
    shapeConfig.Add "SHAPE_POSITION_BK_COLOR", SHAPE_POSITION_BK_COLOR
    shapeConfig.Add "SHAPE_LABEL_TYPE", SHAPE_LABEL_TYPE
    shapeConfig.Add "SHAPE_LABEL_WIDTH", SHAPE_LABEL_WIDTH
    shapeConfig.Add "SHAPE_LABEL_HEIGHT", SHAPE_LABEL_HEIGHT
    shapeConfig.Add "SHAPE_LABEL_BK_COLOR", SHAPE_LABEL_BK_COLOR
    Set GetConfigShapePosition = shapeConfig
End Function

Public Function GetDefaultShapeType() As MsoAutoShapeType
    GetDefaultShapeType = msoShapeRectangle
End Function

Public Function GetTblObject() As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    Set GetTblObject = ws.ListObjects(CONFIG_TABLE_NAME)
End Function



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

