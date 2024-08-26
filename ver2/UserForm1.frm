VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7395
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim wsConfig As Worksheet
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim itm As ListItem
    Dim i As Integer
    Dim colName As Variant
    Dim colIndex As Integer
    Dim selectedColumns As New Collection
    Dim colWidth As Integer

    ' �\������񖼂�Collection�ɒǉ�
    selectedColumns.Add "ID"
    selectedColumns.Add "name"
    selectedColumns.Add "addr"
    selectedColumns.Add "addr_add"
    ' �K�v�ȗ񖼂�ǉ�

    ' ���[�N�V�[�g�ƃe�[�u�����w��
    Set ws = ThisWorkbook.Sheets("map")
    Set tbl = ws.ListObjects("tbl_maps")
    
    ' �ݒ�V�[�g�̊m�F�܂��͍쐬
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("ConfigListView")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        Set wsConfig = ThisWorkbook.Sheets.Add
        wsConfig.Name = "ConfigListView"
    End If

    ' ListView�̐ݒ�
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .MultiSelect = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        ' �R���N�V�������̗񖼂����Ƀw�b�_��ǉ�
        For Each colName In selectedColumns
            colIndex = tbl.ListColumns(colName).index
            .ColumnHeaders.Add , , colName, 100 ' �Œ蕝��ݒ�

            ' �ۑ�����Ă��镝������ΓK�p
            colWidth = wsConfig.Cells(1, colIndex).value
            If colWidth > 0 Then
                .ColumnHeaders(.ColumnHeaders.Count).width = colWidth
            End If
        Next colName
        
        ' �e�[�u���̃f�[�^��ǉ�
        For Each tblRow In tbl.ListRows
            Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, tbl.ListColumns(selectedColumns(1)).index).value)
            For i = 2 To selectedColumns.Count
                itm.SubItems(i - 1) = tblRow.Range.Cells(1, tbl.ListColumns(selectedColumns(i)).index).value
            Next i
        Next tblRow
    End With
End Sub

Private Sub UserForm_Terminate()
    Dim wsConfig As Worksheet
    Dim i As Integer
    
    ' �ݒ�V�[�g���擾
    Set wsConfig = ThisWorkbook.Sheets("ConfigListView")
    
    ' �񕝂�ۑ�
    With Me.ListView1
        For i = 1 To .ColumnHeaders.Count
            wsConfig.Cells(1, i).value = .ColumnHeaders(i).width
        Next i
    End With
End Sub



Private Sub CommandButton1_Click()
    Dim itm As ListItem
    Dim checkedItems As Collection
    Set checkedItems = New Collection

    ' �`�F�b�N����Ă���A�C�e���̃e�L�X�g���R���N�V�����ɒǉ�
    For Each itm In Me.ListView1.ListItems
        If itm.Checked Then
            checkedItems.Add itm.Text
        End If
    Next itm

    ' �R���N�V�����̓��e��Debug.Print�ŏo��
    Dim item As Variant
    For Each item In checkedItems
        Debug.Print item
    Next item
End Sub

