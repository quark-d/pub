Dim nameListA As Collection
Dim nameListB As Collection
Dim nameListC As Collection

' コンボボックスの初期設定を行うプロシージャ
Sub InitializeComboBoxes()
    ' 名前リストの初期化
    Set nameListA = New Collection
    Set nameListB = New Collection
    Set nameListC = New Collection

    ' サンプルデータの追加
    nameListA.Add "ItemA1"
    nameListA.Add "ItemA2"
    nameListA.Add "ItemA3"
    
    nameListB.Add "ItemB1"
    nameListB.Add "ItemB2"
    nameListB.Add "ItemB3"

    nameListC.Add "ItemC1"
    nameListC.Add "ItemC2"
    nameListC.Add "ItemC3"

    ' コンボボックスに項目を追加
    Call PopulateComboBox(cmbBoxA, nameListA)
    Call PopulateComboBox(cmbBoxB, nameListB)
    Call PopulateComboBox(cmbBoxC, nameListC)
    
    ' 初期選択を空白にする
    cmbBoxA.ListIndex = 0
    cmbBoxB.ListIndex = 0
    cmbBoxC.ListIndex = 0
End Sub

' コンボボックスに項目を追加するサブルーチン
Sub PopulateComboBox(cmbBox As ComboBox, nameList As Collection)
    Dim i As Integer
    cmbBox.Clear
    cmbBox.AddItem " " ' 空白を追加
    For i = 1 To nameList.Count
        cmbBox.AddItem nameList(i)
    Next i
End Sub

' コンボボックスAの選択変更イベント
Private Sub cmbBoxA_Change()
    Call TriggerOutput
End Sub

' コンボボックスBの選択変更イベント
Private Sub cmbBoxB_Change()
    Call TriggerOutput
End Sub

' コンボボックスCの選択変更イベント
Private Sub cmbBoxC_Change()
    Call TriggerOutput
End Sub

' 全てのコンボボックスの選択された項目を関数Outputに渡す
Sub TriggerOutput()
    Dim selectedA As String
    Dim selectedB As String
    Dim selectedC As String
    
    selectedA = cmbBoxA.Value
    selectedB = cmbBoxB.Value
    selectedC = cmbBoxC.Value

    ' Output関数を呼び出す
    Call Output(selectedA, selectedB, selectedC)
End Sub

' Output関数 (ここでは、選択された値を即座に表示する例)
Sub Output(selectedA As String, selectedB As String, selectedC As String)
    MsgBox "A: " & selectedA & vbCrLf & "B: " & selectedB & vbCrLf & "C: " & selectedC
End Sub
