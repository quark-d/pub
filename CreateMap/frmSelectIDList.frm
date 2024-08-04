Option Explicit

Private Sub CommandButton1_Click()
    Dim selectedBoardID As String
    
    ' リストボックスで選択されているIDを取得
    If Me.lstBoards.ListIndex <> -1 Then
        selectedBoardID = Me.lstBoards.Value
        ' UpdateShapesVisibilityを呼び出す
        Call UpdateShapesVisibility(selectedBoardID, getIdDict())
        ' フォームを閉じる
        Me.Hide
    Else
        MsgBox "Please select a board ID.", vbExclamation
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim idDict As Object
    Dim boardID As Variant
    
    ' ID辞書を取得する
    Set idDict = getIdDict() ' IDを取得する関数を呼び出す
    
    ' リストボックスにIDを追加
    For Each boardID In idDict.Keys
        ' Me.lstBoards.AddItem boardID
        Me.ListBox1.AddItem boardID
    Next boardID
    
    ' リストボックスのサイズを調整
    Me.ListBox1.Height = Me.Height - 80 ' フォームの高さに合わせて調整
End Sub

