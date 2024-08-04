Option Explicit

Private Sub CommandButton1_Click()
    Dim selectedID As String
    Dim ctrl As Control
    
    ' 選択されたラジオボタンを確認
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" And ctrl.Value = True Then
            selectedID = ctrl.Caption
            Exit For
        End If
    Next ctrl
    
    ' IDが選択されていない場合の処理
    If selectedID = "" Then
        MsgBox "IDが選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' UpdateShapesVisibilityを呼び出す
    Call UpdateShapesVisibility(selectedID, getIdDict())
    
    ' フォームを閉じる
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Dim idDict As Object
    Dim boardID As Variant
    Dim i As Integer
    Dim radioButtonTop As Integer
    Dim radioButtonHeight As Integer
    Dim formHeight As Integer
    Dim optButton As MSForms.OptionButton
    
    ' ID辞書を取得する
    Set idDict = getIdDict() ' IDを取得する関数を呼び出す
    
    ' ラジオボタンの初期化
    i = 1
    radioButtonTop = 20 ' ラジオボタンの上からの位置
    radioButtonHeight = 20 ' ラジオボタンの高さ
    
    ' すべてのラジオボタンを追加
    For Each boardID In idDict.Keys
        ' ラジオボタンを追加する
        Set optButton = Me.Controls.Add("Forms.OptionButton.1", "optBoard" & i, True)
        
        ' ラジオボタンの設定
        optButton.Caption = boardID
        optButton.Visible = True
        optButton.Top = radioButtonTop
        optButton.Left = 20
        
        ' 次のラジオボタンの位置を調整
        radioButtonTop = radioButtonTop + radioButtonHeight + 5
        i = i + 1
    Next boardID
    
    ' フォームの高さを調整
    formHeight = radioButtonTop + 60 ' フォームの高さをラジオボタンの位置 + OKボタンの余白
    Me.Height = formHeight
    
    ' イベント処理を行うために少し待つ
    DoEvents
End Sub





