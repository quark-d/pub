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
'        MsgBox "�ʒu���ۑ�����܂����B"
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
    ' �K�v�ɉ����ăN���[���A�b�v����������
End Sub

Public Sub LoadConfigData(sheetName As String, tableName As String)
    ' �V�[�g���ƃe�[�u������ private �ϐ��ɕۑ�
    CONFIG_SHEET_NAME = sheetName
    CONFIG_TABLE_NAME = tableName

    Dim ws As Worksheet
    Dim configTable As ListObject
    
    ' �V�[�g�����݂��Ȃ��ꍇ�͍쐬
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = CONFIG_SHEET_NAME
    End If
    
    ' �e�[�u�������݂��Ȃ��ꍇ�͍쐬
    On Error Resume Next
    Set configTable = ws.ListObjects(CONFIG_TABLE_NAME)
    On Error GoTo 0
    If configTable Is Nothing Then
        ' �w�b�_�[�s���Z�b�g
        ws.Range("A1:F1").value = Array("ShapeID", "X", "Y", "RGB_R", "RGB_G", "RGB_B")
        ' �e�[�u�����쐬
        Set configTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F1"), , xlYes)
        configTable.Name = CONFIG_TABLE_NAME
    End If
    
    ' �e�[�u������ݒ�f�[�^��ǂݍ���
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

    ' �V�[�g�ƃe�[�u�����擾
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    Set configTable = ws.ListObjects(CONFIG_TABLE_NAME)
    
    ' �e�[�u�����N���A
    configTable.DataBodyRange.ClearContents
    
    ' ��������f�[�^���e�[�u���ɏ����߂�
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
    ' �ʒu����ǉ��܂��͍X�V����
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

    ' �V�[�g�ƈʒu���𕪊����Ď擾
    sheetName = Split(positionString, DELIMITER_ADDR)(0)
    position = Split(positionString, DELIMITER_ADDR)(1)
    shapeName = "pos_" & sheetName & "_" & position

    ' �ʒu���V�[�g
    Set posWs = ThisWorkbook.Sheets(CONFIG_SHEET_NAME)
    ' �e�[�u��
    Set tbl = posWs.ListObjects(CONFIG_TABLE_NAME)
    
    ' �Y���V�[�g�̊m�F�܂��͍쐬
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    End If
    On Error GoTo 0

    ' �K���RGB�F���擾
    On Error Resume Next
    ' �_�~�[�V�F�C�v���쐬���ċK��̐F���擾
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 10, 10)
    colorRGB = shp.Fill.ForeColor.RGB
    redValue = colorRGB And 255
    greenValue = (colorRGB \ 256) And 255
    blueValue = (colorRGB \ 65536) And 255
    shp.Delete ' �_�~�[�V�F�C�v���폜
    On Error GoTo 0

    ' �����̃��R�[�h������
    Dim foundRow As ListRow
    For Each foundRow In tbl.ListRows
        If foundRow.Range(1).value = shapeName Then Exit Sub
    Next foundRow

    ' �V�K���R�[�h��ǉ�
    With tbl.ListRows.Add
        .Range(1).value = shapeName
        .Range(2).value = 0 ' ����X���W
        .Range(3).value = 0 ' ����Y���W
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
    
    ' ConfigPositions �e�[�u������ ShapeID ������
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
    ' Collection�ɓ���̒l�����݂��邩���m�F����w���p�[�֐�
    Dim item As Variant
    For Each item In col
        If item = value Then
            ExistsInCollection = True
            Exit Function
        End If
    Next item
    ExistsInCollection = False
End Function

