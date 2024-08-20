Attribute VB_Name = "modDrowShapes"
Option Explicit

Private Const DELIMITER_ADDR As String = "/"
Private Const DELIMITER_ADDRS As String = ":"



'Sub DrowPosition(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = "/")
'    Dim ws As Worksheet
'    Dim shapeName As String
'    Dim sheetName As String
'    Dim position As String
'    Dim shapeID As String
'    Dim shape As shape
'    Dim xPos As Double
'    Dim yPos As Double
'    Dim redValue As Integer
'    Dim greenValue As Integer
'    Dim blueValue As Integer
'    Dim shapeDict As Dictionary
'    Dim existingShapes As Collection
'    Dim shp As shape
'
'    ' �ʒu���𕪉�
'    sheetName = Split(positionString, delimiter)(0)
'    position = Split(positionString, delimiter)(1)
'    shapeID = "pos_" & sheetName & "_" & position
'
'    ' ������ "pos_" �v���t�B�b�N�X�̃I�[�g�V�F�C�v���X�g���擾
'    Set existingShapes = GetAutoShapeList("pos_")
'
'    ' ConfigPositions ������W�ƐF�̏����擾
'    configPosition.AllocateShapeAttribute shapeID, xPos, yPos, redValue, greenValue, blueValue
'
'    ' ���łɕ`�悳��Ă��邩���`�F�b�N
'    For Each shp In existingShapes
'        If shp.Name = shapeID Then
'            ' �w�i�F�̂ݕύX
'            shp.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
'            Exit Sub ' �������I��
'        End If
'    Next shp
'
'    ' �Y���V�[�g�̊m�F�܂��͍쐬
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets(sheetName)
'    If ws Is Nothing Then
'        Set ws = ThisWorkbook.Sheets.Add
'        ws.Name = sheetName
'    End If
'    On Error GoTo 0
'
'    ' ConfigShapes �̐ݒ���擾
'    Set shapeDict = configShape.GetConfigShapePosition()
'
'    ' �I�[�g�V�F�C�v�̕`��
'    Set shape = ws.Shapes.AddShape(shapeDict("SHAPE_POSITION_TYPE"), xPos, yPos, shapeDict("SHAPE_POSITION_WIDTH"), shapeDict("SHAPE_POSITION_HEIGHT"))
'    shape.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
'    shape.Name = shapeID ' �V�F�C�v����ݒ�
'    shape.TextFrame2.TextRange.Text = position
'End Sub
'Public Sub DrowShapes(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = "/")
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim positionList As Collection
'    Dim row As ListRow
'    Dim positions() As String
'    Dim individualPosition As Variant
'    Dim positionParts() As String
'    Dim i As Integer
'
'    Call DrowStandby(configShape, configPosition, positionList)
'
'    ' �V�[�g�ƃe�[�u���̐ݒ�
'    Set ws = ThisWorkbook.Sheets("map")
'    Set tbl = ws.ListObjects("tbl_maps")
'
'    ' �e�[�u���̊e�s�����[�v
'    For Each row In tbl.ListRows
'        ' addr �񂩂�ʒu�����擾
'        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, ":")
'
'        ' ���ꂼ��̈ʒu�����[�v���� DrowPosition ���Ăяo��
'        For i = LBound(positions) To UBound(positions)
'            individualPosition = positions(i)
'            ' DrowPosition ���Ăяo��
'            Call DrowShape(configShape, configPosition, CStr(individualPosition))
'        Next i
'    Next row
'
'End Sub

Public Sub DrowShapes(configShape As ConfigShapes, configPosition As ConfigPositions, positionList As Collection, Optional delimiter As String = DELIMITER_ADDR)
    Dim individualPosition As Variant
    
    ' �����֐��̌Ăяo��
    Call DrowStandby(configShape, configPosition, positionList)
    
    ' �ʒu���X�g�����[�v���� DrowShape ���Ăяo��
    For Each individualPosition In positionList
        ' DrowShape ���Ăяo��
        Call DrowShape(configShape, configPosition, CStr(individualPosition), delimiter)
    Next individualPosition
End Sub


'Private Sub DrowShape(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = "/")
'    Dim ws As Worksheet
'    Dim labelName As String
'    Dim sheetName As String
'    Dim position As String
'    Dim shapeID As String
'    Dim xPos As Double
'    Dim yPos As Double
'    Dim redValue As Integer
'    Dim greenValue As Integer
'    Dim blueValue As Integer
'    Dim shapeDict As Dictionary
'    Dim existingShapes As Collection
'    Dim existingShape As shape
'    Dim shp As shape
'
'    ' �ʒu���𕪉�
'    sheetName = Split(positionString, delimiter)(0)
'    position = Split(positionString, delimiter)(1)
'    shapeID = "pos_" & sheetName & "_" & position
'    labelName = shapeID & "_lbl"
'
'    ' ������ "pos_" �v���t�B�b�N�X�̃I�[�g�V�F�C�v���X�g���擾
'    Set existingShapes = GetAutoShapeListPosition()
'
'    ' ConfigPositions ������W�ƐF�̏����擾
'    configPosition.AllocateShapeAttribute shapeID, xPos, yPos, redValue, greenValue, blueValue
'
'
'    ' �Y���V�[�g�̊m�F�܂��͍쐬
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets(sheetName)
'    If ws Is Nothing Then
'        Set ws = ThisWorkbook.Sheets.Add
'        ws.Name = sheetName
'    End If
'    On Error GoTo 0
'
'    ' ConfigShapes �̐ݒ���擾
'    Set shapeDict = configShape.GetConfigShapePosition()
'
'
'    ' ���łɕ`�悳��Ă��邩���`�F�b�N
'    Set existingShape = Nothing
'    For Each shp In existingShapes
'        If shp.Name = shapeID Then
'            Set existingShape = shp
'            ' �w�i�F�̂ݕύX
'            shp.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
'
'            ' ���x���ƃR�l�N�^��`��
'            On Error Resume Next ' �G���[������ǉ�
'            Call DrowLabel(ws, shapeDict, labelName, shp.left, shp.top)
'            Call DrowConnection(ws, shp, ws.Shapes(labelName))
'            On Error GoTo 0
'
'            Exit Sub ' �������I��
'        End If
'    Next shp
'
'          ' position��label��`�悵�A�����Ō���
'          Call DrowPosition(ws, shapeDict, shapeID, xPos, yPos, redValue, greenValue, blueValue, position)
'          Call DrowLabel(ws, shapeDict, labelName, xPos, yPos)
'          Call DrowConnection(ws, ws.Shapes(shapeID), ws.Shapes(labelName))
'
'End Sub
Private Sub DrowShape(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = DELIMITER_ADDR)
    Dim ws As Worksheet
    Dim labelName As String
    Dim sheetName As String
    Dim position As String
    Dim shapeID As String
    Dim xPos As Double
    Dim yPos As Double
    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer
    Dim shapeDict As Dictionary
    Dim existingShapes As Collection
    Dim existingShape As shape
    Dim shp As shape
    Dim labelShapes As Collection
    Dim existingLabel As shape

    ' �ʒu���𕪉�
    sheetName = Split(positionString, delimiter)(0)
    position = Split(positionString, delimiter)(1)
    shapeID = "pos_" & sheetName & "_" & position
    labelName = shapeID & "_lbl"

    ' ������ "pos_" �v���t�B�b�N�X�̃I�[�g�V�F�C�v���X�g���擾
    Set existingShapes = GetAutoShapeListPosition()
    
    ' ConfigPositions ������W�ƐF�̏����擾
    configPosition.AllocateShapeAttribute shapeID, xPos, yPos, redValue, greenValue, blueValue

    ' �Y���V�[�g�̊m�F�܂��͍쐬
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    End If
    On Error GoTo 0

    ' ConfigShapes �̐ݒ���擾
    Set shapeDict = configShape.GetConfigShapePosition()
    
    ' ���łɕ`�悳��Ă��邩���`�F�b�N
    Set existingShape = Nothing
    For Each shp In existingShapes
        If shp.Name = shapeID Then
            Set existingShape = shp
            ' �w�i�F�̂ݕύX
            shp.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
            
            ' ���x���V�F�C�v�̑��݂��m�F
            Set labelShapes = GetAutoShapeListLabel()
            Set existingLabel = Nothing
            For Each existingLabel In labelShapes
                If existingLabel.Name = labelName Then
                    Exit Sub ' ���x�������݂��邽�ߏ������I��
                End If
            Next existingLabel
            
            ' ���x���ƃR�l�N�^��`��
            Call DrowLabel(ws, shapeDict, labelName, shp.left, shp.top)
            Call DrowConnection(ws, shp, ws.Shapes(labelName))
            
            Exit Sub ' �������I��
        End If
    Next shp
          
    ' position��label��`�悵�A�����Ō���
    Call DrowPosition(ws, shapeDict, shapeID, xPos, yPos, redValue, greenValue, blueValue, position)
    Call DrowLabel(ws, shapeDict, labelName, xPos, yPos)
    Call DrowConnection(ws, ws.Shapes(shapeID), ws.Shapes(labelName))

End Sub



Private Sub DrowPosition(ws As Worksheet, shapeDict As Dictionary, shapeID As String, xPos As Double, yPos As Double, redValue As Integer, greenValue As Integer, blueValue As Integer, position As String)
    Dim shape As shape
    ' position�I�[�g�V�F�C�v�̕`��
    Set shape = ws.Shapes.AddShape(shapeDict("SHAPE_POSITION_TYPE"), xPos, yPos, shapeDict("SHAPE_POSITION_WIDTH"), shapeDict("SHAPE_POSITION_HEIGHT"))
    shape.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
    shape.Name = shapeID ' �V�F�C�v����ݒ�
    shape.TextFrame2.TextRange.Text = position
End Sub
Private Sub DrowLabel(ws As Worksheet, shapeDict As Dictionary, labelName As String, xPos As Double, yPos As Double)
    Dim label As shape
    Dim labelType As MsoAutoShapeType
    'Dim labelType As String
    Dim left As Single
    Dim top As Single
    Dim width As Single
    Dim height As Single

    'labelType = shapeDict("SHAPE_LABEL_TYPE")
    left = xPos
    top = yPos + shapeDict("SHAPE_POSITION_HEIGHT") + 10
    width = shapeDict("SHAPE_LABEL_WIDTH")
    height = shapeDict("SHAPE_LABEL_HEIGHT")
    
    ' label�I�[�g�V�F�C�v�̕`��iposition�̉��ɕ\���j
    Set label = ws.Shapes.AddShape(shapeDict("SHAPE_LABEL_TYPE"), left, top, width, height)
    label.Fill.ForeColor.RGB = RGB(200, 200, 200) ' �����͔����D�F
    label.Name = labelName
    label.TextFrame2.TextRange.Text = "" ' �󕶎�
    label.Line.Visible = msoFalse ' �O���̐����\��
End Sub

Private Sub DrowConnection(ws As Worksheet, shape As shape, label As shape)
    Dim connector As shape
    ' position��label�𒼐��Ō���
    Set connector = ws.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
    connector.ConnectorFormat.BeginConnect shape, 2 ' position�V�F�C�v�̒�����
    connector.ConnectorFormat.EndConnect label, 1   ' label�V�F�C�v�̒�����
    connector.Line.Weight = 1.5 ' ���̑���
End Sub

Private Function GetAutoShapeListPosition() As Collection
    Dim shapeList As New Collection
    Dim ws As Worksheet
    Dim shp As shape
    
    ' �S�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g���̂��ׂẴV�F�C�v�����[�v
        For Each shp In ws.Shapes
            ' "pos_" �v���t�B�b�N�X�Ŏn�܂�A"_lbl" �T�t�B�b�N�X�ŏI���Ȃ��ꍇ�A���X�g�ɒǉ�
            If shp.Name Like "pos_*" And Not shp.Name Like "*_lbl" Then
                shapeList.Add shp
            End If
        Next shp
    Next ws
    
    ' ���ʂ�Ԃ�
    Set GetAutoShapeListPosition = shapeList
End Function

Public Function GetAutoShapeListLabel() As Collection
    Dim shapeList As New Collection
    Dim ws As Worksheet
    Dim shp As shape
    
    ' �S�V�[�g�����[�v
    For Each ws In ThisWorkbook.Worksheets
        ' �V�[�g���̂��ׂẴV�F�C�v�����[�v
        For Each shp In ws.Shapes
            ' "pos_" �v���t�B�b�N�X�Ŏn�܂�A"_lbl" �T�t�B�b�N�X�ŏI���ꍇ�A���X�g�ɒǉ�
            If shp.Name Like "pos_*_lbl" Then
                shapeList.Add shp
            End If
        Next shp
    Next ws
    
    ' ���ʂ�Ԃ�
    Set GetAutoShapeListLabel = shapeList
End Function


Private Sub DrowStandby(configShape As ConfigShapes, configPosition As ConfigPositions, positionList As Collection)
  Dim existingShapes As Collection
  Dim shp As shape
  Dim labelName As String
  
  ' "pos_" ����n�܂邷�ׂẴV�F�C�v���擾
  Set existingShapes = GetAutoShapeListPosition()
  
  ' ���ׂĂ� "pos_" ����n�܂�V�F�C�v�����[�v
  For Each shp In existingShapes
      ' ���x���p�̃T�t�B�b�N�X�����V�F�C�v�����쐬
      labelName = shp.Name & "_lbl"
  
      ' ���x���I�[�g�V�F�C�v���폜
      On Error Resume Next
      shp.Parent.Shapes(labelName).Delete
      On Error GoTo 0
  
      ' �ڑ��V�F�C�v���폜
'      Dim connector As shape
'      For Each connector In shp.Parent.Shapes
'          If connector.Type = MsoConnector Then
'              If connector.ConnectorFormat.BeginConnectedShape Is shp Or connector.ConnectorFormat.EndConnectedShape Is shp Then
'                  connector.Delete
'              End If
'          End If
'      Next connector
'      Dim connector As shape
'      For Each connector In shp.Parent.Shapes
'          If connector.connector Then
'              ' �V�F�C�v���R�l�N�^�ł���A���݂̃V�F�C�v�ɐڑ�����Ă��邩���`�F�b�N
'              If connector.ConnectorFormat.BeginConnectedShape Is shp Or connector.ConnectorFormat.EndConnectedShape Is shp Then
'                  connector.Delete
'              End If
'          End If
'      Next connector
        Dim connector As shape
        For Each connector In shp.Parent.Shapes
            ' �V�F�C�v���R�l�N�^���ǂ������m�F
            If connector.connector Then
                ' �R�l�N�^�����݂̃V�F�C�v�ɐڑ�����Ă��邩���`�F�b�N
                If Not connector.ConnectorFormat.BeginConnectedShape Is Nothing Then
                    If connector.ConnectorFormat.BeginConnectedShape Is shp Then
                        connector.Delete
                    End If
'                    If connector.ConnectorFormat.EndConnectedShape Is shp Then
'                      connector.Delete
'                    End If
                End If
            End If
        Next connector
  
      ' �c�����V�F�C�v�𔖂��D�F�ɐݒ�
      shp.Fill.ForeColor.RGB = RGB(200, 200, 200)
  Next shp

End Sub
Sub AppendTextInLabel(tbl As ListObject, configShape As ConfigShapes, selectedIDs As Collection)
    Dim ws As Worksheet
    Dim row As ListRow
    Dim positions() As String
    Dim individualPosition As Variant
    Dim sheetName As String
    Dim position As String
    Dim shapeID As String
    Dim labelName As String
    Dim labelShape As shape
    Dim posShape As shape
    Dim shapeDict As Dictionary
    Dim xPos As Double
    Dim yPos As Double
    Dim labelShapes As Collection
    Dim shp As shape
    
    ' ConfigShapes �̐ݒ���擾
    Set shapeDict = configShape.GetConfigShapePosition()
    
    ' �e�[�u���̊e�s�����[�v
    For Each row In tbl.ListRows
        ' addr �񂩂�ʒu�����擾
        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, DELIMITER_ADDRS)
        
        ' ���ꂼ��̈ʒu�����[�v
        For Each individualPosition In positions
            sheetName = Split(individualPosition, DELIMITER_ADDR)(0)
            position = Split(individualPosition, DELIMITER_ADDR)(1)
            shapeID = "pos_" & sheetName & "_" & position
            labelName = shapeID & "_lbl"

            ' �V�[�g�̎擾�܂��͍쐬
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(sheetName)
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets.Add
                ws.Name = sheetName
            End If
            On Error GoTo 0

            ' �����̃��x���V�F�C�v���X�g���擾
            Set labelShapes = GetAutoShapeListLabel()
            
            ' �����̃��x���V�F�C�v��T��
            Set labelShape = Nothing
            For Each shp In labelShapes
                If shp.Name = labelName Then
                    Set labelShape = shp
                    Exit For
                End If
            Next shp

            ' ���x���V�F�C�v��������Ȃ��ꍇ�A�V�K�쐬���X�L�b�v
            If labelShape Is Nothing Then
                ' ���̈ʒu���ɃX�L�b�v
                GoTo ContinueLoop
            End If

            ' selectedIDs �ƈ�v���邩�m�F
            Dim selectedID As Variant
            For Each selectedID In selectedIDs
                If row.Range(1).value = selectedID Then
                    ' ���x���V�F�C�v�Ƀe�L�X�g��ǉ�
                    Call SetSheetTabColor(labelShape.Parent)
                    If Not labelShape Is Nothing Then
                        If labelShape.TextFrame2.TextRange.Text <> "" Then
                            labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & vbCrLf
                        End If
                        labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & row.Range(1).value
                    End If
                    Exit For
                End If
            Next selectedID

ContinueLoop:
        Next individualPosition
    Next row
End Sub

'Sub AppendTextInLabel(tbl As ListObject, configShape As ConfigShapes)
'    Dim ws As Worksheet
'    Dim row As ListRow
'    Dim positions() As String
'    Dim individualPosition As Variant
'    Dim sheetName As String
'    Dim position As String
'    Dim shapeID As String
'    Dim labelName As String
'    Dim labelShape As shape
'    Dim posShape As shape
'    Dim shapeDict As Dictionary
'    Dim xPos As Double
'    Dim yPos As Double
'    Dim labelShapes As Collection
'    Dim shp As shape
'
'    ' ConfigShapes �̐ݒ���擾
'    Set shapeDict = configShape.GetConfigShapePosition()
'
'    ' �e�[�u���̊e�s�����[�v
'    For Each row In tbl.ListRows
'        ' addr �񂩂�ʒu�����擾
'        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, ":")
'
'        ' ���ꂼ��̈ʒu�����[�v
'        For Each individualPosition In positions
'            sheetName = Split(individualPosition, "/")(0)
'            position = Split(individualPosition, "/")(1)
'            shapeID = "pos_" & sheetName & "_" & position
'            labelName = shapeID & "_lbl"
'
'            ' �V�[�g�̎擾�܂��͍쐬
'            On Error Resume Next
'            Set ws = ThisWorkbook.Sheets(sheetName)
'            If ws Is Nothing Then
'                Set ws = ThisWorkbook.Sheets.Add
'                ws.Name = sheetName
'            End If
'            On Error GoTo 0
'
'            ' �����̃��x���V�F�C�v���X�g���擾
'            Set labelShapes = GetAutoShapeListLabel()
'
'            ' �����̃��x���V�F�C�v��T��
'            Set labelShape = Nothing
'            For Each shp In labelShapes
'                If shp.Name = labelName Then
'                    Set labelShape = shp
'                    Exit For
'                End If
'            Next shp
'
'            ' ���x���V�F�C�v��������Ȃ��ꍇ�A�V�K�쐬
'            If labelShape Is Nothing Then
'                ' Position�V�F�C�v�̈ʒu���擾
'                On Error Resume Next
'                Set posShape = ws.Shapes(shapeID)
'                On Error GoTo 0
'
'                If Not posShape Is Nothing Then
'                    ' Position�V�F�C�v�̉��Ƀ��x�����쐬
'                    xPos = posShape.left
'                    yPos = posShape.top + posShape.height + 10
'                    Set labelShape = ws.Shapes.AddShape(shapeDict("SHAPE_LABEL_TYPE"), xPos, yPos, shapeDict("SHAPE_LABEL_WIDTH"), shapeDict("SHAPE_LABEL_HEIGHT"))
'                    labelShape.Name = labelName
'                    labelShape.Fill.ForeColor.RGB = RGB(200, 200, 200)
'                    labelShape.TextFrame2.TextRange.Text = ""
'                    labelShape.Line.Visible = msoFalse ' �O���̐����\��
'                End If
'            End If
'
'            ' ���x���V�F�C�v�Ƀe�L�X�g��ǉ�
'            If Not labelShape Is Nothing Then
'                If labelShape.TextFrame2.TextRange.Text <> "" Then
'                    labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & vbCrLf
'                End If
'                labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & row.Range(1).value
'            End If
'        Next individualPosition
'    Next row
'End Sub
'
'Public Sub AdjustShapeLabelByText(labelShape As shape)
'    Dim textLines() As String
'    Dim maxWidth As Double
'    Dim i As Integer
'
'    ' �����s�̃e�L�X�g���l�����čł������s�̕����v�Z
'    textLines = Split(labelShape.TextFrame2.TextRange.Text, vbCrLf)
'    maxWidth = 0
'    For i = LBound(textLines) To UBound(textLines)
'        labelShape.TextFrame2.TextRange.Text = textLines(i)
'        If labelShape.TextFrame2.TextRange.BoundWidth > maxWidth Then
'            maxWidth = labelShape.TextFrame2.TextRange.BoundWidth
'        End If
'    Next i
'
'    ' �e�L�X�g�̍����ƍł������s�ɍ��킹�ă��x���V�F�C�v�̃T�C�Y�𒲐�
'    labelShape.height = labelShape.TextFrame2.TextRange.BoundHeight + 10 ' �����𒲐�
'    labelShape.width = maxWidth + 10   ' �ł������s�Ɋ�Â��ĕ��𒲐�
'End Sub
Sub AdjustShapeLabelByText(labelShape As shape)
    Dim textLines() As String
    Dim maxWidth As Double
    Dim i As Integer
    
    ' �����܂�Ԃ��𖳌���
    labelShape.TextFrame2.WordWrap = msoFalse
    
    ' �����s�̃e�L�X�g���l�����čł������s�̕����v�Z
    textLines = Split(labelShape.TextFrame2.TextRange.Text, vbCrLf)
    maxWidth = 0
    For i = LBound(textLines) To UBound(textLines)
        labelShape.TextFrame2.TextRange.Text = textLines(i)
        If labelShape.TextFrame2.TextRange.BoundWidth > maxWidth Then
            maxWidth = labelShape.TextFrame2.TextRange.BoundWidth
        End If
    Next i
    
    ' �e�L�X�g�̍����ƍł������s�ɍ��킹�ă��x���V�F�C�v�̃T�C�Y�𒲐�
    labelShape.height = labelShape.TextFrame2.TextRange.BoundHeight + 10 ' �����𒲐�
    'labelShape.width = maxWidth + 10   ' �ł������s�Ɋ�Â��ĕ��𒲐�
    
    ' �����܂�Ԃ����ēx�L�����i�K�v�ɉ����āj
    labelShape.TextFrame2.WordWrap = msoTrue
End Sub

