VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYesNoCancel 
   Caption         =   "frmYesNoCancel"
   ClientHeight    =   465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "frmYesNoCancel.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmYesNoCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private result As VbMsgBoxResult

Private Sub btnYes_Click()
    result = vbYes
    Me.Hide
End Sub

Private Sub btnNo_Click()
    result = vbNo
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    result = vbCancel
    Me.Hide
End Sub

Public Function ShowDialog(Optional ByVal formCaption As String = "", _
                           Optional ByVal yesCaption As String = "Yes", _
                           Optional ByVal noCaption As String = "No", _
                           Optional ByVal cancelCaption As String = "Cancel") As VbMsgBoxResult
    Me.Caption = formCaption
    Me.btnYes.Caption = yesCaption
    Me.btnNo.Caption = noCaption
    Me.btnCancel.Caption = cancelCaption
    
    Me.Show vbModal
    ShowDialog = result
End Function

'�g�p��
'Sub ChangeShapeColor()
'    Dim shpName As String
'    Dim shp As shape
'    Dim userChoice As VbMsgBoxResult
'
'    ' Application.Caller ���g���ăN���b�N���ꂽ�V�F�C�v�̖��O���擾
'    shpName = Application.Caller
'
'    ' �V�F�C�v�I�u�W�F�N�g���擾
'    Set shp = ActiveSheet.Shapes(shpName)
'
'    ' frmYesNoCancel�t�H�[����\�����A�I�����ꂽ�{�^���̌��ʂ��擾
'    userChoice = frmYesNoCancel.ShowDialog("�w�i�F�ύX", "�ǉ�", "�폜", "�L�����Z��")
'
'    ' �I���ɉ����Ĕw�i�F��ύX
'    Select Case userChoice
'        Case vbYes ' �ǉ���I��
'            shp.Fill.ForeColor.RGB = RGB(255, 255, 0) ' ���F
'
'        Case vbNo ' �폜��I��
'            shp.Fill.ForeColor.RGB = RGB(255, 0, 0) ' ��
'
'        Case vbCancel ' �L�����Z����I��
'            ' �������Ȃ�
'    End Select
'End Sub


'---------------------------------------------------------------
''ActiveSheet�̂��ׂẴV�F�C�v�Ƀ}�N�����蓖��
'---------------------------------------------------------------
'Sub AssignMacroToAllShapes()
'    Dim shp As shape
'
'    ' �A�N�e�B�u�V�[�g��̂��ׂẴV�F�C�v�ɑ΂��ă��[�v
'    For Each shp In ActiveSheet.Shapes
'        ' �V�F�C�v�� "ChangeShapeColor" �}�N�������蓖��
'        shp.OnAction = "ChangeShapeColor"
'    Next shp
'End Sub


''����̃V�F�C�v�ɑ΂��ă}�N���ݒ�
'Sub AssignMacroToPositionShapes()
'    Dim shp As shape
'    Dim shapeList As Collection
'    Dim shpItem As Variant
'
'    ' ����̃V�F�C�v���܂ރR���N�V�������擾
'    Set shapeList = GetAutoShapeListPosition()
'
'    ' �擾�����V�F�C�v���X�g���̊e�V�F�C�v�ɑ΂��ă��[�v
'    For Each shpItem In shapeList
'        ' �e�V�F�C�v�� "ChangeShapeColor" �}�N�������蓖��
'        shpItem.OnAction = "ChangeShapeColor"
'    Next shpItem
'End Sub

'---------------------------------------------------------------
'ID,�@���l�A���l��SharePoint�փA�b�v�f�[�g
'---------------------------------------------------------------
'Sub UpdateSharePointRecords(recordCollection As Collection)
'    ' �ϐ��錾
'    Dim conn As Object ' ADODB.Connection
'    Dim rs As Object ' ADODB.Recordset
'    Dim strConn As String
'    Dim sharePointURL As String
'    Dim listName As String
'    Dim wsConfig As Worksheet
'    Dim tblConfig As ListObject
'    Dim record As Object
'    Dim i As Integer
'
'    ' configConnect�V�[�g��tbl_config_con�e�[�u���̊m�F
'    On Error Resume Next
'    Set wsConfig = ThisWorkbook.Sheets("configConnect")
'    On Error GoTo 0
'
'    ' �V�[�g�����݂��Ȃ��ꍇ�͍쐬
'    If wsConfig Is Nothing Then
'        Set wsConfig = ThisWorkbook.Sheets.Add
'        wsConfig.Name = "configConnect"
'        MsgBox "configConnect�V�[�g���쐬����܂����BURL�ƃ��X�g�����w�肵�Ă��������B", vbInformation
'    End If
'
'    ' �e�[�u���̊m�F
'    On Error Resume Next
'    Set tblConfig = wsConfig.ListObjects("tbl_config_con")
'    On Error GoTo 0
'
'    ' �e�[�u�������݂��Ȃ��ꍇ�͍쐬
'    If tblConfig Is Nothing Then
'        Set tblConfig = wsConfig.ListObjects.Add(xlSrcRange, wsConfig.Range("A1:B2"), , xlYes)
'        tblConfig.Name = "tbl_config_con"
'        wsConfig.Range("A1").value = "url"
'        wsConfig.Range("B1").value = "listName"
'        MsgBox "tbl_config_con�e�[�u�����쐬����܂����BURL�ƃ��X�g�����w�肵�Ă��������B", vbInformation
'        Exit Sub
'    End If
'
'    ' URL�ƃ��X�g���̎擾
'    sharePointURL = tblConfig.DataBodyRange.Cells(1, 1).value
'    listName = tblConfig.DataBodyRange.Cells(1, 2).value
'
'    ' URL�܂��̓��X�g������̏ꍇ�̓��b�Z�[�W��\�����ďI��
'    If sharePointURL = "" Or listName = "" Then
'        MsgBox "URL�܂��̓��X�g�����w�肳��Ă��܂���B", vbCritical
'        Exit Sub
'    End If
'
'    ' �ڑ�������̍쐬
'    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'              "WSS;IMEX=1;RetrieveIds=Yes;" & _
'              "DATABASE=" & sharePointURL
'
'    ' �ڑ��I�u�W�F�N�g�̍쐬
'    Set conn = CreateObject("ADODB.Connection")
'    conn.Open strConn
'
'    ' RecordSet�I�u�W�F�N�g�̍쐬
'    Set rs = CreateObject("ADODB.Recordset")
'
'    ' Collection���̊e���R�[�h������
'    For i = 1 To recordCollection.Count
'        Set record = recordCollection(i)
'
'        ' ���R�[�h�Z�b�g�̑���
'        With rs
'            ' �����ID_ID�������R�[�h���擾
'            .Open "SELECT * FROM [" & listName & "] WHERE ID_ID = '" & record("ID_ID") & "'", conn, 1, 3
'
'            ' ���R�[�h�����݂��邩�m�F
'            If Not .EOF Then
'                .Fields("COL1").value = record("COL1")
'                .Fields("COL2").value = record("COL2")
'                .Update ' ���R�[�h���X�V
'            Else
'                MsgBox "�w�肳�ꂽID_ID (" & record("ID_ID") & ") �̃��R�[�h��������܂���B"
'            End If
'            .Close ' RecordSet�����i���̃��[�v�ɔ����āj
'        End With
'    Next i
'
'    ' RecordSet��Connection�����
'    conn.Close
'    Set rs = Nothing
'    Set conn = Nothing
'
'    MsgBox "���ׂẴ��R�[�h���X�V���܂����B"
'End Sub

'
'Sub ExampleUsage()
'    ' Collection���쐬���ēn��
'    Dim recordCollection As Collection
'    Set recordCollection = CreateCollectionOfRecords
'
'    ' Collection�������Ƃ��ēn���čX�V���������s
'    Call UpdateSharePointRecords(recordCollection)
'End Sub
'
'Function CreateCollectionOfRecords() As Collection
'    ' ��Ƃ��Đ������ꂽCollection���g�p
'    Dim recordCollection As New Collection
'    Dim record As Object
'    Dim i As Integer
'
'    For i = 1 To 5
'        Set record = CreateObject("Scripting.Dictionary")
'        record("ID_ID") = "ID_" & CStr(i)
'        record("COL1") = i * 10.5
'        record("COL2") = i * 20.8
'        recordCollection.Add record
'    Next i
'
'    Set CreateCollectionOfRecords = recordCollection
'End Function


'---------------------------------------------------------------
'2 �̃e�[�u��a , b�̤�w�肵������r����قȂ���̂�ID���擾����
'---------------------------------------------------------------
'' �V�[�g���̒�`
'Private Const QUERY_SHEET_A As String = "SheetA"
'Private Const QUERY_SHEET_B As String = "SheetB"
'
'' �e�[�u�����̒�`
'Private Const QUERY_TABLE_A As String = "tbl_a"
'Private Const QUERY_TABLE_B As String = "tbl_b"
'
'' �񖼂̒�` (QUERY_TABLE_A)
'Private Const QUERY_TABLE_A_DOC As String = "doc"
'Private Const QUERY_TABLE_A_DOCREV As String = "docrev"
'
'' �񖼂̒�` (QUERY_TABLE_B)
'Private Const QUERY_TABLE_B_DOC As String = "doc"
'Private Const QUERY_TABLE_B_DOCREV As String = "docrev"
'Private Const QUERY_TABLE_B_DOCID As String = "docID"
'
'Sub UpdateTablesAndCompare()
'    ' �ϐ��錾
'    Dim wsA As Worksheet, wsB As Worksheet
'    Dim tblA As ListObject, tblB As ListObject
'
'    ' �e�[�u����ێ����Ă���V�[�g���w��
'    Set wsA = ThisWorkbook.Sheets(QUERY_SHEET_A) ' tbl_a������V�[�g
'    Set wsB = ThisWorkbook.Sheets(QUERY_SHEET_B) ' tbl_b������V�[�g
'
'    ' �e�[�u���I�u�W�F�N�g���擾
'    Set tblA = wsA.ListObjects(QUERY_TABLE_A)
'    Set tblB = wsB.ListObjects(QUERY_TABLE_B)
'
'    ' �e�[�u�����X�V
'    tblA.QueryTable.Refresh BackgroundQuery:=False
'    tblB.QueryTable.Refresh BackgroundQuery:=False
'
'    ' ��r���s���T�u���[�`�����Ăяo��
'    GetDifferentDocIDsUsingDictionary tblA, tblB
'End Sub
'
'Sub GetDifferentDocIDsUsingDictionary(tblA As ListObject, tblB As ListObject)
'    ' �ϐ��錾
'    Dim docDictionary As Object
'    Dim colDocIDs As New Collection
'    Dim rowA As ListRow, rowB As ListRow
'    Dim docKey As String
'    Dim docrevA As Long, docrevB As Long
'    Dim resultDict As Object
'
'    ' Dictionary�̍쐬
'    Set docDictionary = CreateObject("Scripting.Dictionary")
'
'    ' tbl_a�̃f�[�^��Dictionary�ɕێ� (doc���L�[�Adocrev��l�Ƃ���)
'    For Each rowA In tblA.ListRows
'        docKey = rowA.Range.Cells(1, tblA.ListColumns(QUERY_TABLE_A_DOC).index).value
'        docrevA = rowA.Range.Cells(1, tblA.ListColumns(QUERY_TABLE_A_DOCREV).index).value
'        docDictionary(docKey) = docrevA
'    Next rowA
'
'    ' tbl_b�����[�v���āADictionary�Ƃ̔�r���s��
'    For Each rowB In tblB.ListRows
'        docKey = rowB.Range.Cells(1, tblB.ListColumns(QUERY_TABLE_B_DOC).index).value
'        docrevB = rowB.Range.Cells(1, tblB.ListColumns(QUERY_TABLE_B_DOCREV).index).value
'
'        ' docKey����łȂ����m�F
'        If docKey <> "" Then
'            ' Dictionary��doc�����݂��Adocrev���قȂ�Atbl_a��docrev���������ꍇ
'            If docDictionary.Exists(docKey) Then
'                If docDictionary(docKey) < docrevB Then
'                    ' �V����Dictionary���쐬
'                    Set resultDict = CreateObject("Scripting.Dictionary")
'                    resultDict("docID") = rowB.Range.Cells(1, tblB.ListColumns(QUERY_TABLE_B_DOCID).index).value
'                    resultDict("docrevChange") = docDictionary(docKey) & " --> " & docrevB
'                    ' Collection�ɒǉ�
'                    colDocIDs.Add resultDict
'                End If
'            End If
'        End If
'    Next rowB
'
'    ' ���ʂ��m�F
'    Dim docEntry As Variant
'    For Each docEntry In colDocIDs
'        Debug.Print "docID: " & docEntry("docID") & ", docrevChange: " & docEntry("docrevChange")
'    Next docEntry
'End Sub

