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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)
    
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
    positions = Split(addr, DELIMITER_ADDRS)
    
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
        positions = Split(row.Range(tbl.ListColumns(QUERY_ITEM_ADDR).Index).value, DELIMITER_ADDRS)
        
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
    Set ws = ThisWorkbook.Sheets(QUERY_SHEET)
    Set tbl = ws.ListObjects(QUERY_TABLE)

    ' �e�[�u������ID��addr�̃y�A���擾���Ď����ɕۑ�
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
    
    ' �������̊e�A�C�e���𔽕�����
    For Each item In idAddrDict.Items
        ' addr������� ":" �ŕ���
        addrParts = Split(item, ":")
        
        ' �e�����𔽕��������āA�V�[�g���𒊏o
        For i = LBound(addrParts) To UBound(addrParts)
            classroomParts = Split(addrParts(i), "/")
            sheetName = classroomParts(0)
            
            ' �R���N�V�������ŃV�[�g�������ɑ��݂��邩�m�F
            alreadyExists = False
            For Each existingSheet In sheetNames
                If existingSheet = sheetName Then
                    alreadyExists = True
                    Exit For
                End If
            Next existingSheet
            
            ' �V�[�g�������j�[�N�ł���΃R���N�V�����ɒǉ�
            If Not alreadyExists Then
                sheetNames.Add sheetName
            End If
        Next i
    Next item
    
    ' �V�[�g���̃R���N�V������Ԃ�
    Set GetSheetNamesFromAddrDict = sheetNames
End Function


Public Sub ResetAllSheetTabColors()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Tab.colorIndex = xlColorIndexNone
    Next ws
End Sub
Public Sub SetSheetTabColor(ws As Worksheet, Optional colorIndex As Long = 3)
    ' �w�肳�ꂽ�V�[�g�̃^�u�̐F�� colorIndex �Őݒ�
    ws.Tab.colorIndex = colorIndex
End Sub

