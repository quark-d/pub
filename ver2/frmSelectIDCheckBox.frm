VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectIDCheckBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectIDCheckBox.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    
    ' 選択されたチェックボックスを確認
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" And ctrl.value = True Then
            selectedItems.Add ctrl.Caption
        End If
    Next ctrl
    
    ' チェックされた項目がない場合の処理
    If selectedItems.Count = 0 Then
        MsgBox "項目が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' Collectionに格納された選択項目の処理
    Dim item As Variant
    For Each item In selectedItems
        Debug.Print item ' ここで選択された項目を使用する処理を行う
    Next item
    
     Call EntryPointMulti(selectedItems)
    
    ' フォームを閉じる
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
    
    ' ID辞書を取得する
    Set idAddrDict = GetIdAddrDictionary() ' IDを取得する関数を呼び出す
    
    ' チェックボックスの初期化
    i = 1
    checkBoxTop = 20 ' チェックボックスの上からの位置
    checkBoxHeight = 20 ' チェックボックスの高さ
    
    ' すべてのチェックボックスを追加
    For Each boardID In idAddrDict.Keys
        ' チェックボックスを追加する
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkBox" & i, True)
        
        ' チェックボックスの設定
        chkBox.Caption = boardID
        chkBox.Visible = True
        chkBox.top = checkBoxTop
        chkBox.left = 20
        
        ' 次のチェックボックスの位置を調整
        checkBoxTop = checkBoxTop + checkBoxHeight + 5
        i = i + 1
    Next boardID
    
    ' フォームの高さを調整
    formHeight = checkBoxTop + 60 ' フォームの高さをチェックボックスの位置 + OKボタンの余白
    Me.height = formHeight
    
    ' イベント処理を行うために少し待つ
    DoEvents
End Sub

