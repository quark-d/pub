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
    
    ' �V�[�g�ƃe�[�u���̐ݒ�
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    
    ' �ʒu���X�g���擾
    Set positionList = GetLocationList(tbl)
    
    For Each individualPosition In positionList
        configPosition.SetUpPosition CStr(individualPosition)
    Next individualPosition
    
    ' DrowPositions �֐����Ăяo��
    Call DrowShapes(configShape, configPosition, positionList)
    
    Call AppendTextInLabel(tbl, configShape)

    ' ���x���̃T�C�Y�𒲐�
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
    
    ' �e�[�u���̊e�s�����[�v
    For Each row In tbl.ListRows
        ' addr �񂩂�ʒu�����擾
        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, ":")
        
        ' ���ꂼ��̈ʒu�����[�v���ăR���N�V�����ɒǉ�
        For i = LBound(positions) To UBound(positions)
            individualPosition = positions(i)
            alreadyExists = False
            
            ' �R���N�V�������Ɋ��ɑ��݂��邩�`�F�b�N
            For Each item In positionList
                If item = individualPosition Then
                    alreadyExists = True
                    Exit For
                End If
            Next item
            
            ' ���ɑ��݂��Ȃ���΃R���N�V�����ɒǉ�
            If Not alreadyExists Then
                positionList.Add individualPosition
            End If
        Next i
    Next row
    
    ' �R���N�V������Ԃ�
    Set GetLocationList = positionList
End Function

