VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectIDRadio 
   Caption         =   "ID��I�����Ă��������B"
   ClientHeight    =   1005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   OleObjectBlob   =   "frmSelectIDRadio.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSelectIDRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim selectedID As String
    Dim ctrl As Control
    
    ' �I�����ꂽ���W�I�{�^�����m�F
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" And ctrl.value = True Then
            selectedID = ctrl.Caption
            Exit For
        End If
    Next ctrl
    
    ' ID���I������Ă��Ȃ��ꍇ�̏���
    If selectedID = "" Then
        MsgBox "ID���I������Ă��܂���B", vbExclamation
        Exit Sub
    End If
    
    Call ResetAllSheetTabColors
    ' UpdateShapesVisibility���Ăяo��
    ' Call UpdateShapesVisibility(selectedID, getIdDict())
    Call EntryPointOne(selectedID)
    
    
    ' �t�H�[�������
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Dim idDict As Object
    Dim idAddrDict As Scripting.Dictionary
    Dim boardID As Variant
    Dim i As Integer
    Dim radioButtonTop As Integer
    Dim radioButtonHeight As Integer
    Dim formHeight As Integer
    Dim optButton As MSForms.OptionButton
    
    ' ID�������擾����
    ' Set idDict = getIdDict() ' ID���擾����֐����Ăяo��
    Set idAddrDict = GetIdAddrDictionary() ' ID���擾����֐����Ăяo��
    
    ' ���W�I�{�^���̏�����
    i = 1
    radioButtonTop = 20 ' ���W�I�{�^���̏ォ��̈ʒu
    radioButtonHeight = 20 ' ���W�I�{�^���̍���
    
    ' ���ׂẴ��W�I�{�^����ǉ�
    For Each boardID In idAddrDict.Keys
        ' ���W�I�{�^����ǉ�����
        Set optButton = Me.Controls.Add("Forms.OptionButton.1", "optBoard" & i, True)
        
        ' ���W�I�{�^���̐ݒ�
        optButton.Caption = boardID
        optButton.Visible = True
        optButton.top = radioButtonTop
        optButton.left = 20
        
        ' ���̃��W�I�{�^���̈ʒu�𒲐�
        radioButtonTop = radioButtonTop + radioButtonHeight + 5
        i = i + 1
    Next boardID
    
    ' �t�H�[���̍����𒲐�
    formHeight = radioButtonTop + 60 ' �t�H�[���̍��������W�I�{�^���̈ʒu + OK�{�^���̗]��
    Me.height = formHeight
    
    ' �C�x���g�������s�����߂ɏ����҂�
    DoEvents
End Sub





