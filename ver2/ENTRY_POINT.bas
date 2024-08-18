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
    
    ' configPosition�̏�����
    Call InitializeconfigPosition
    
    Dim idAddrDict As Scripting.Dictionary
    Set idAddrDict = GetIdAddrDictionary() ' ID���擾����֐����Ăяo��
    
    ' ������selectedID������
    For Each selectedID In selectedIDs
        If Not idAddrDict.Exists(selectedID) Then
            MsgBox "�w�肳�ꂽID " & selectedID & " �͎����ɑ��݂��܂���B", vbExclamation
            Exit Sub
        End If
        
        Dim addr As String
        addr = idAddrDict(selectedID)
        
        ' �ʒu���X�g���擾���ApositionList�ɒǉ�
        Dim tempPositionList As Collection
        Set tempPositionList = GetLocationListSingle(addr)
        
        For Each individualPosition In tempPositionList
            positionList.Add individualPosition
        Next individualPosition
    Next selectedID
    
    ' �ʒu���Z�b�g�A�b�v
    For Each individualPosition In positionList
        configPosition.SetUpPosition CStr(individualPosition)
    Next individualPosition
    
    ' DrowPositions �֐����Ăяo��
    Call DrowShapes(configShape, configPosition, positionList)
    
    ' �V�[�g�ƃe�[�u���̐ݒ�
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    Call AppendTextInLabel(tbl, configShape, selectedIDs)

    ' ���x���̃T�C�Y�𒲐�
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
    Set idAddrDict = GetIdAddrDictionary() ' ID���擾����֐����Ăяo��
    If Not idAddrDict.Exists(selectedID) Then
        MsgBox "�w�肳�ꂽselectedID�͎����ɑ��݂��܂���B", vbExclamation
        Exit Sub
    End If
    
    Dim addr As String
    addr = idAddrDict(selectedID)
    ' �ʒu���X�g���擾
    Set positionList = GetLocationListSingle(addr)
    
    For Each individualPosition In positionList
        configPosition.SetUpPosition CStr(individualPosition)
    Next individualPosition
    
    ' DrowPositions �֐����Ăяo��
    Call DrowShapes(configShape, configPosition, positionList)
    
    ' �V�[�g�ƃe�[�u���̐ݒ�
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    Dim selectedIDs As New Collection ' Collection���쐬
    selectedIDs.Add selectedID
    Call AppendTextInLabel(tbl, configShape, selectedIDs)

    ' ���x���̃T�C�Y�𒲐�
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
    
    Dim row As ListRow
    Dim idCollection As New Collection ' ID���i�[����Collection���쐬
    ' �e�[�u������ID��̒l���擾���ACollection�ɒǉ�
    For Each row In tbl.ListRows
        idCollection.Add row.Range(1).value ' ID���܂܂�Ă������w��i�����ł�1��ځj
    Next row
    
    Call AppendTextInLabel(tbl, configShape, idCollection)

    ' ���x���̃T�C�Y�𒲐�
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
    
    ' �ʒu���� ":" �ŕ���
    positions = Split(addr, ":")
    
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
    
    ' �R���N�V������Ԃ�
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

Function GetIdAddrDictionary() As Scripting.Dictionary
    Dim idAddrDict As New Scripting.Dictionary
    Dim row As ListRow
    Dim id As String
    Dim addr As String
    
    ' �V�[�g�ƃe�[�u���̐ݒ�
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")

    ' �e�[�u������ID��addr�̃y�A���擾���Ď����ɕۑ�
    For Each row In tbl.ListRows
        id = row.Range(tbl.ListColumns("ID").Index).value
        addr = row.Range(tbl.ListColumns("addr").Index).value
        idAddrDict.Add id, addr
    Next row
    
    Set GetIdAddrDictionary = idAddrDict
End Function
