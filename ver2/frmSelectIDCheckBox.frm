VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectIDCheckBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectIDCheckBox.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSelectIDCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim selectedItems As New Collection
    Dim ctrl As Control
    
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
    Dim item As Variant
    For Each item In selectedItems
        Debug.Print item ' �����őI�����ꂽ���ڂ��g�p���鏈�����s��
    Next item
    
     Call EntryPointMulti(selectedItems)
    
    ' �t�H�[�������
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Dim idAddrDict As Scripting.Dictionary
    Dim boardID As Variant
    Dim i As Integer
    Dim checkBoxTop As Integer
    Dim checkBoxHeight As Integer
    Dim formHeight As Integer
    Dim chkBox As MSForms.CheckBox
    
    ' ID�������擾����
    Set idAddrDict = GetIdAddrDictionary() ' ID���擾����֐����Ăяo��
    
    ' �`�F�b�N�{�b�N�X�̏�����
    i = 1
    checkBoxTop = 20 ' �`�F�b�N�{�b�N�X�̏ォ��̈ʒu
    checkBoxHeight = 20 ' �`�F�b�N�{�b�N�X�̍���
    
    ' ���ׂẴ`�F�b�N�{�b�N�X��ǉ�
    For Each boardID In idAddrDict.Keys
        ' �`�F�b�N�{�b�N�X��ǉ�����
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkBox" & i, True)
        
        ' �`�F�b�N�{�b�N�X�̐ݒ�
        chkBox.Caption = boardID
        chkBox.Visible = True
        chkBox.top = checkBoxTop
        chkBox.left = 20
        
        ' ���̃`�F�b�N�{�b�N�X�̈ʒu�𒲐�
        checkBoxTop = checkBoxTop + checkBoxHeight + 5
        i = i + 1
    Next boardID
    
    ' �t�H�[���̍����𒲐�
    formHeight = checkBoxTop + 60 ' �t�H�[���̍������`�F�b�N�{�b�N�X�̈ʒu + OK�{�^���̗]��
    Me.height = formHeight
    
    ' �C�x���g�������s�����߂ɏ����҂�
    DoEvents
End Sub

