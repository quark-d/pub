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

'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim header As String
'    Dim i As Long
'    Dim dataRow As ListRow
'
'    ' �V�[�g�ƃe�[�u�����擾
'    Set ws = ThisWorkbook.Sheets("map") ' �V�[�g�����w��
'    Set tbl = ws.ListObjects("tbl_maps")
'
'    ' �w�b�_��ListBox�ɒǉ�
'    header = ""
'    For i = 1 To tbl.HeaderRowRange.Columns.Count
'        If i > 1 Then header = header & vbTab ' �^�u�ŋ�؂�
'        header = header & tbl.HeaderRowRange.Cells(1, i).value
'    Next i
'    Me.ListBox1.AddItem header
'
'    ' �e�[�u���̃f�[�^��ListBox�ɒǉ�
'    For Each dataRow In tbl.ListRows
'        rowData = ""
'        For i = 1 To dataRow.Range.Columns.Count
'            If i > 1 Then rowData = rowData & vbTab ' �^�u�ŋ�؂�
'            rowData = rowData & dataRow.Range.Cells(1, i).value
'        Next i
'        Me.ListBox1.AddItem rowData
'    Next dataRow
'End Sub


'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim dataRow As ListRow
'    Dim i As Long
'    Dim lbl As MSForms.label
'    Dim leftPos As Single
'    Dim columnWidth As Single
'    Dim currentWidth As Single
'
'    ' �V�[�g�ƃe�[�u�����擾
'    Set ws = ThisWorkbook.Sheets("map") ' �V�[�g�����w��
'    Set tbl = ws.ListObjects("tbl_maps")
'
'    ' �񕝂𓮓I�Ɍv�Z (ListBox�̕���񐔂Ŋ���)
'    columnWidth = Me.ListBox1.width / tbl.HeaderRowRange.Columns.Count
'
'    ' �����̍��ʒu��ݒ�
'    leftPos = Me.ListBox1.left
'
'    ' �w�b�_��\�����x���𓮓I�ɐ���
'    For i = 1 To tbl.HeaderRowRange.Columns.Count
'        Set lbl = Me.Controls.Add("Forms.Label.1")
'
'        With lbl
'            .Caption = tbl.HeaderRowRange.Cells(1, i).value
'            .width = columnWidth ' �v�Z�����񕝂�ݒ�
'            ' ��̒����Ƀ��x����z�u
'            .left = leftPos + (columnWidth - .width) / 2
'            .top = Me.ListBox1.top - 15
'            .TextAlign = fmTextAlignCenter
'        End With
'
'        ' ���̃��x���̍��ʒu�𒲐�
'        leftPos = leftPos + columnWidth
'    Next i
'
'    ' ListBox�̐ݒ�
'    With Me.ListBox1
'        .ColumnCount = tbl.HeaderRowRange.Columns.Count ' �e�[�u���̗񐔂�ݒ�
'        .BoundColumn = 1
'        .ColumnHeads = False
'    End With
'
'    ' �f�[�^��ListBox�ɒǉ�����
'    For Each dataRow In tbl.ListRows
'        Dim rowData(1 To 10) As String
'        For i = 1 To tbl.HeaderRowRange.Columns.Count
'            rowData(i) = dataRow.Range.Cells(1, i).value
'        Next i
'        Me.ListBox1.AddItem Join(rowData, vbTab)
'    Next dataRow
'End Sub
''
'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRow As ListRow
'    Dim dataArray() As String
'    Dim i As Long
'
'    ' ���[�N�V�[�g�ƃe�[�u�����w��
'    Set ws = ThisWorkbook.Sheets("tbl_maps")
'    Set tbl = ws.ListObjects("tbl_maps_1")
'
'    ' �w�b�_��ǉ�
'    With Me.ListBox1
'        .ColumnCount = tbl.ListColumns.Count
'        .AddItem "ID" & vbTab & "name" & vbTab & "addr"
'
'        ' �f�[�^��ǉ�
'        For Each tblRow In tbl.ListRows
'            ReDim dataArray(1 To tbl.ListColumns.Count)
'            For i = 1 To tbl.ListColumns.Count
'                dataArray(i) = tblRow.Range(1, i).value
'            Next i
'            .AddItem Join(dataArray, vbTab)
'        Next tblRow
'    End With
'End Sub

'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRow As ListRow
'    Dim itm As ListItem
'    Dim i As Integer
'    Dim colWidth As Integer
'
'    ' ���[�N�V�[�g�ƃe�[�u�����w��
'    Set ws = ThisWorkbook.Sheets("maps")
'    Set tbl = ws.ListObjects("tbl_maps_1")
'
'    ' ListView�̐ݒ�
'    With Me.ListView1
'        .View = lvwReport
'        .Gridlines = True
'        .FullRowSelect = True
'        .ColumnHeaders.Clear
'        .ListItems.Clear
'
'        ' �e�[�u���̃w�b�_��ǉ�
'        For i = 1 To tbl.ListColumns.Count
'            .ColumnHeaders.Add , , tbl.ListColumns(i).Name, 100 ' ��������ݒ�
'        Next i
'
'        ' �e�[�u���̃f�[�^��ǉ�
'        For Each tblRow In tbl.ListRows
'            Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, 1).value) ' �ŏ��̗�̃f�[�^
'            For i = 2 To tbl.ListColumns.Count
'                itm.SubItems(i - 1) = tblRow.Range.Cells(1, i).value
'            Next i
'        Next tblRow
'
'        ' �񕝂��蓮�Œ���
'        For i = 1 To .ColumnHeaders.Count
'            colWidth = .ColumnHeaders(i).width
'            ' �f�[�^�ƃw�b�_�̒����Ɋ�Â��ė񕝂�ݒ�
'            For Each itm In .ListItems
'                colWidth = Application.Max(colWidth, Me.TextWidth(itm.Text) + 10)
'            Next itm
'            .ColumnHeaders(i).width = colWidth
'        Next i
'    End With
'End Sub


'Dim savedSelections As Collection
'
'Private Sub UserForm_Initialize()
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tblRow As ListRow
'    Dim itm As ListItem
'    Dim i As Integer
'
'    ' ���[�N�V�[�g�ƃe�[�u�����w��
'    Set ws = ThisWorkbook.Sheets("tbl_maps")
'    Set tbl = ws.ListObjects("tbl_maps_1")
'
'    ' �I����Ԃ�ۑ����邽�߂̃R���N�V������������
'    Set savedSelections = New Collection
'
'    ' ListView�̐ݒ�
'    With Me.ListView1
'        .View = lvwReport
'        .Gridlines = True
'        .FullRowSelect = True
'        .MultiSelect = True ' �����I����L���ɂ���
'        .ColumnHeaders.Clear
'        .ListItems.Clear
'
'        ' �e�[�u���̃w�b�_��ǉ�
'        For i = 1 To tbl.ListColumns.Count
'            .ColumnHeaders.Add , , tbl.ListColumns(i).Name, 100 ' �Œ蕝��ݒ�
'        Next i
'
'        ' �e�[�u���̃f�[�^��ǉ�
'        For Each tblRow In tbl.ListRows
'            Set itm = .ListItems.Add(, , tblRow.Range.Cells(1, 1).value) ' �ŏ��̗�̃f�[�^
'            For i = 2 To tbl.ListColumns.Count
'                itm.SubItems(i - 1) = tblRow.Range.Cells(1, i).value
'            Next i
'        Next tblRow
'    End With
'End Sub
'
'Private Sub UserForm_Deactivate()
'    ' �I�����ꂽ�A�C�e����ۑ�
'    Dim itm As ListItem
'    Set savedSelections = New Collection
'    For Each itm In Me.ListView1.selectedItems
'        savedSelections.Add itm.Text
'    Next itm
'End Sub
'
'Private Sub UserForm_Activate()
'    ' �ۑ����ꂽ�I���𕜌�
'    Dim itm As ListItem
'    Dim selItem As Variant
'
'    If savedSelections.Count > 0 Then
'        For Each selItem In savedSelections
'            For Each itm In Me.ListView1.ListItems
'                If itm.Text = selItem Then
'                    itm.Selected = True
'                    Exit For
'                End If
'            Next itm
'        Next selItem
'    End If
'End Sub

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

