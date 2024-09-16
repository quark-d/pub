' ComboBoxA の MouseDown イベント
Private Sub cmbBoxA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' 左クリック (Button = 1) 右クリック(Button=2) または中央ボタン (Button = 4) で空文字を設定
    If Button = 2 Or Button = 4 Then
        cmbBoxA.value = "" ' 空文字を設定
        ' Call TriggerOutput ' Output関数を呼び出す
    End If
End Sub


' リセットボタンが押された時に、cmbBoxAの選択を空文字にリセットしOutput関数を呼び出す
Private Sub btnResetA_Click()
    cmbBoxA.Value = "" ' 空文字を選択
    Call TriggerOutput ' Output関数を呼び出す
End Sub

' リセットボタンが押された時に、cmbBoxBの選択を空文字にリセットしOutput関数を呼び出す
Private Sub btnResetB_Click()
    cmbBoxB.Value = "" ' 空文字を選択
    Call TriggerOutput ' Output関数を呼び出す
End Sub

' リセットボタンが押された時に、cmbBoxCの選択を空文字にリセットしOutput関数を呼び出す
Private Sub btnResetC_Click()
    cmbBoxC.Value = "" ' 空文字を選択
    Call TriggerOutput ' Output関数を呼び出す
End Sub

' Output関数を呼び出すために各ComboBoxの値を取得し、Output関数を呼び出す
Sub TriggerOutput()
    Dim selectedA As String
    Dim selectedB As String
    Dim selectedC As String
    
    ' 各コンボボックスの選択された値を取得
    selectedA = cmbBoxA.Value
    selectedB = cmbBoxB.Value
    selectedC = cmbBoxC.Value

    ' Output関数を呼び出す
    Call Output(selectedA, selectedB, selectedC)
End Sub

' Output関数 (例として選択された値を表示するだけの処理)
Sub Output(selectedA As String, selectedB As String, selectedC As String)
    MsgBox "A: " & selectedA & vbCrLf & "B: " & selectedB & vbCrLf & "C: " & selectedC
End Sub
