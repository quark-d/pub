VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYesNo 
   Caption         =   "frmYesNo"
   ClientHeight    =   480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2820
   OleObjectBlob   =   "frmYesNo.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmYesNo"
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


Public Function ShowDialog(Optional ByVal formCaption As String = "", _
                           Optional ByVal yesCaption As String = "Yes", _
                           Optional ByVal noCaption As String = "No") As VbMsgBoxResult
    Me.Caption = formCaption
    Me.btnYes.Caption = yesCaption
    Me.btnNo.Caption = noCaption
    
    Me.Show vbModal
    ShowDialog = result
End Function


