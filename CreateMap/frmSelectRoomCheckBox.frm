VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectRoomCheckBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectRoomCheckBox.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSelectRoomCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim selectedItems As New Collection
    Dim ctrl As Control
    Dim ws As Worksheet
    
    ' �I�����ꂽ�`�F�b�N�{�b�N�X���m�F
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" And ctrl.value = True Then
            selectedItems.Add ctrl.Caption
        End If
    Next ctrl
    
    ' �`�F�b�N���ꂽ���ڂ��Ȃ��ꍇ�̏���
    If selectedItems.Count = 0 Then
        MsgBox "���ڂ��I������Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
   
    ' Collection�Ɋi�[���ꂽ�I�����ڂ̏���
    Call ResetAllSheetTabColors
    
    Dim item As Variant
    For Each item In selectedItems
        ' �Y������V�[�g�̃^�u�� SetSheetTabColor �Őɂ��� (ColorIndex = 5)
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = item Then
                Call SetSheetTabColor(ws, 5) ' �F�ɐݒ�
                Exit For
            End If
        Next ws
    Next item
    
    ' �F�ɐݒ肳�ꂽ�V�[�g�݂̂����
    Call ExportSheetsWithColorTabsToPDF(5)
    
    ' �t�H�[�������
    Me.Hide
End Sub



Private Sub UserForm_Initialize()
    Dim idAddrDict As Scripting.Dictionary
    Dim shtNameCollection As Collection
    
    Dim shtName As Variant
    Dim topPosition As Integer
    Dim chkBox As MSForms.CheckBox
    Dim formHeight As Integer
    Dim spacing As Integer
    
    ' �`�F�b�N�{�b�N�X�̊Ԋu
    spacing = 20
    
    ' ID�������擾����
    Set idAddrDict = GetIdAddrDictionary() ' ID���擾����֐����Ăяo��
    Set shtNameCollection = GetSheetNamesFromAddrDict(idAddrDict)
    
    ' �`�F�b�N�{�b�N�X�̔z�u�J�n�ʒu
    topPosition = 10
    
    ' �V�[�g�����ƂɃ`�F�b�N�{�b�N�X���쐬
    For Each shtName In shtNameCollection
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1")
        With chkBox
            .Caption = shtName
            .left = 10
            .top = topPosition
            .width = 150
            topPosition = topPosition + spacing
        End With
    Next shtName
    
    ' �t�H�[���̍����𒲐�
    formHeight = topPosition + 40 ' �`�F�b�N�{�b�N�X�̍Ō�̈ʒu�ɗ]����ǉ�
    Me.height = formHeight
    
    ' �C�x���g�������s�����߂ɏ����҂�
    DoEvents
End Sub


