VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmYesNoCancel 
   Caption         =   "frmYesNoCancel"
   ClientHeight    =   465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "frmYesNoCancel.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmYesNoCancel"
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

Private Sub btnCancel_Click()
    result = vbCancel
    Me.Hide
End Sub

Public Function ShowDialog(Optional ByVal formCaption As String = "", _
                           Optional ByVal yesCaption As String = "Yes", _
                           Optional ByVal noCaption As String = "No", _
                           Optional ByVal cancelCaption As String = "Cancel") As VbMsgBoxResult
    Me.Caption = formCaption
    Me.btnYes.Caption = yesCaption
    Me.btnNo.Caption = noCaption
    Me.btnCancel.Caption = cancelCaption
    
    Me.Show vbModal
    ShowDialog = result
End Function

'使用例
'Sub ChangeShapeColor()
'    Dim shpName As String
'    Dim shp As shape
'    Dim userChoice As VbMsgBoxResult
'
'    ' Application.Caller を使ってクリックされたシェイプの名前を取得
'    shpName = Application.Caller
'
'    ' シェイプオブジェクトを取得
'    Set shp = ActiveSheet.Shapes(shpName)
'
'    ' frmYesNoCancelフォームを表示し、選択されたボタンの結果を取得
'    userChoice = frmYesNoCancel.ShowDialog("背景色変更", "追加", "削除", "キャンセル")
'
'    ' 選択に応じて背景色を変更
'    Select Case userChoice
'        Case vbYes ' 追加を選択
'            shp.Fill.ForeColor.RGB = RGB(255, 255, 0) ' 黄色
'
'        Case vbNo ' 削除を選択
'            shp.Fill.ForeColor.RGB = RGB(255, 0, 0) ' 赤
'
'        Case vbCancel ' キャンセルを選択
'            ' 何もしない
'    End Select
'End Sub


'---------------------------------------------------------------
''ActiveSheetのすべてのシェイプにマクロ割り当て
'---------------------------------------------------------------
'Sub AssignMacroToAllShapes()
'    Dim shp As shape
'
'    ' アクティブシート上のすべてのシェイプに対してループ
'    For Each shp In ActiveSheet.Shapes
'        ' シェイプに "ChangeShapeColor" マクロを割り当て
'        shp.OnAction = "ChangeShapeColor"
'    Next shp
'End Sub


''特定のシェイプに対してマクロ設定
'Sub AssignMacroToPositionShapes()
'    Dim shp As shape
'    Dim shapeList As Collection
'    Dim shpItem As Variant
'
'    ' 特定のシェイプを含むコレクションを取得
'    Set shapeList = GetAutoShapeListPosition()
'
'    ' 取得したシェイプリスト内の各シェイプに対してループ
'    For Each shpItem In shapeList
'        ' 各シェイプに "ChangeShapeColor" マクロを割り当て
'        shpItem.OnAction = "ChangeShapeColor"
'    Next shpItem
'End Sub

'---------------------------------------------------------------
'ID,　数値、数値をSharePointへアップデート
'---------------------------------------------------------------
'Sub UpdateSharePointRecords(recordCollection As Collection)
'    ' 変数宣言
'    Dim conn As Object ' ADODB.Connection
'    Dim rs As Object ' ADODB.Recordset
'    Dim strConn As String
'    Dim sharePointURL As String
'    Dim listName As String
'    Dim wsConfig As Worksheet
'    Dim tblConfig As ListObject
'    Dim record As Object
'    Dim i As Integer
'
'    ' configConnectシートとtbl_config_conテーブルの確認
'    On Error Resume Next
'    Set wsConfig = ThisWorkbook.Sheets("configConnect")
'    On Error GoTo 0
'
'    ' シートが存在しない場合は作成
'    If wsConfig Is Nothing Then
'        Set wsConfig = ThisWorkbook.Sheets.Add
'        wsConfig.Name = "configConnect"
'        MsgBox "configConnectシートが作成されました。URLとリスト名を指定してください。", vbInformation
'    End If
'
'    ' テーブルの確認
'    On Error Resume Next
'    Set tblConfig = wsConfig.ListObjects("tbl_config_con")
'    On Error GoTo 0
'
'    ' テーブルが存在しない場合は作成
'    If tblConfig Is Nothing Then
'        Set tblConfig = wsConfig.ListObjects.Add(xlSrcRange, wsConfig.Range("A1:B2"), , xlYes)
'        tblConfig.Name = "tbl_config_con"
'        wsConfig.Range("A1").value = "url"
'        wsConfig.Range("B1").value = "listName"
'        MsgBox "tbl_config_conテーブルが作成されました。URLとリスト名を指定してください。", vbInformation
'        Exit Sub
'    End If
'
'    ' URLとリスト名の取得
'    sharePointURL = tblConfig.DataBodyRange.Cells(1, 1).value
'    listName = tblConfig.DataBodyRange.Cells(1, 2).value
'
'    ' URLまたはリスト名が空の場合はメッセージを表示して終了
'    If sharePointURL = "" Or listName = "" Then
'        MsgBox "URLまたはリスト名が指定されていません。", vbCritical
'        Exit Sub
'    End If
'
'    ' 接続文字列の作成
'    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'              "WSS;IMEX=1;RetrieveIds=Yes;" & _
'              "DATABASE=" & sharePointURL
'
'    ' 接続オブジェクトの作成
'    Set conn = CreateObject("ADODB.Connection")
'    conn.Open strConn
'
'    ' RecordSetオブジェクトの作成
'    Set rs = CreateObject("ADODB.Recordset")
'
'    ' Collection内の各レコードを処理
'    For i = 1 To recordCollection.Count
'        Set record = recordCollection(i)
'
'        ' レコードセットの操作
'        With rs
'            ' 特定のID_IDを持つレコードを取得
'            .Open "SELECT * FROM [" & listName & "] WHERE ID_ID = '" & record("ID_ID") & "'", conn, 1, 3
'
'            ' レコードが存在するか確認
'            If Not .EOF Then
'                .Fields("COL1").value = record("COL1")
'                .Fields("COL2").value = record("COL2")
'                .Update ' レコードを更新
'            Else
'                MsgBox "指定されたID_ID (" & record("ID_ID") & ") のレコードが見つかりません。"
'            End If
'            .Close ' RecordSetを閉じる（次のループに備えて）
'        End With
'    Next i
'
'    ' RecordSetとConnectionを閉じる
'    conn.Close
'    Set rs = Nothing
'    Set conn = Nothing
'
'    MsgBox "すべてのレコードを更新しました。"
'End Sub

'
'Sub ExampleUsage()
'    ' Collectionを作成して渡す
'    Dim recordCollection As Collection
'    Set recordCollection = CreateCollectionOfRecords
'
'    ' Collectionを引数として渡して更新処理を実行
'    Call UpdateSharePointRecords(recordCollection)
'End Sub
'
'Function CreateCollectionOfRecords() As Collection
'    ' 例として生成されたCollectionを使用
'    Dim recordCollection As New Collection
'    Dim record As Object
'    Dim i As Integer
'
'    For i = 1 To 5
'        Set record = CreateObject("Scripting.Dictionary")
'        record("ID_ID") = "ID_" & CStr(i)
'        record("COL1") = i * 10.5
'        record("COL2") = i * 20.8
'        recordCollection.Add record
'    Next i
'
'    Set CreateCollectionOfRecords = recordCollection
'End Function


'---------------------------------------------------------------
'2 つのテーブルa , bの､指定した列を比較し､異なるもののIDを取得する
'---------------------------------------------------------------
'' シート名の定義
'Private Const QUERY_SHEET_A As String = "SheetA"
'Private Const QUERY_SHEET_B As String = "SheetB"
'
'' テーブル名の定義
'Private Const QUERY_TABLE_A As String = "tbl_a"
'Private Const QUERY_TABLE_B As String = "tbl_b"
'
'' 列名の定義 (QUERY_TABLE_A)
'Private Const QUERY_TABLE_A_DOC As String = "doc"
'Private Const QUERY_TABLE_A_DOCREV As String = "docrev"
'
'' 列名の定義 (QUERY_TABLE_B)
'Private Const QUERY_TABLE_B_DOC As String = "doc"
'Private Const QUERY_TABLE_B_DOCREV As String = "docrev"
'Private Const QUERY_TABLE_B_DOCID As String = "docID"
'
'Sub UpdateTablesAndCompare()
'    ' 変数宣言
'    Dim wsA As Worksheet, wsB As Worksheet
'    Dim tblA As ListObject, tblB As ListObject
'
'    ' テーブルを保持しているシートを指定
'    Set wsA = ThisWorkbook.Sheets(QUERY_SHEET_A) ' tbl_aがあるシート
'    Set wsB = ThisWorkbook.Sheets(QUERY_SHEET_B) ' tbl_bがあるシート
'
'    ' テーブルオブジェクトを取得
'    Set tblA = wsA.ListObjects(QUERY_TABLE_A)
'    Set tblB = wsB.ListObjects(QUERY_TABLE_B)
'
'    ' テーブルを更新
'    tblA.QueryTable.Refresh BackgroundQuery:=False
'    tblB.QueryTable.Refresh BackgroundQuery:=False
'
'    ' 比較を行うサブルーチンを呼び出し
'    GetDifferentDocIDsUsingDictionary tblA, tblB
'End Sub
'
'Sub GetDifferentDocIDsUsingDictionary(tblA As ListObject, tblB As ListObject)
'    ' 変数宣言
'    Dim docDictionary As Object
'    Dim colDocIDs As New Collection
'    Dim rowA As ListRow, rowB As ListRow
'    Dim docKey As String
'    Dim docrevA As Long, docrevB As Long
'    Dim resultDict As Object
'
'    ' Dictionaryの作成
'    Set docDictionary = CreateObject("Scripting.Dictionary")
'
'    ' tbl_aのデータをDictionaryに保持 (docをキー、docrevを値として)
'    For Each rowA In tblA.ListRows
'        docKey = rowA.Range.Cells(1, tblA.ListColumns(QUERY_TABLE_A_DOC).index).value
'        docrevA = rowA.Range.Cells(1, tblA.ListColumns(QUERY_TABLE_A_DOCREV).index).value
'        docDictionary(docKey) = docrevA
'    Next rowA
'
'    ' tbl_bをループして、Dictionaryとの比較を行う
'    For Each rowB In tblB.ListRows
'        docKey = rowB.Range.Cells(1, tblB.ListColumns(QUERY_TABLE_B_DOC).index).value
'        docrevB = rowB.Range.Cells(1, tblB.ListColumns(QUERY_TABLE_B_DOCREV).index).value
'
'        ' docKeyが空でないか確認
'        If docKey <> "" Then
'            ' Dictionaryにdocが存在し、docrevが異なり、tbl_aのdocrevが小さい場合
'            If docDictionary.Exists(docKey) Then
'                If docDictionary(docKey) < docrevB Then
'                    ' 新しいDictionaryを作成
'                    Set resultDict = CreateObject("Scripting.Dictionary")
'                    resultDict("docID") = rowB.Range.Cells(1, tblB.ListColumns(QUERY_TABLE_B_DOCID).index).value
'                    resultDict("docrevChange") = docDictionary(docKey) & " --> " & docrevB
'                    ' Collectionに追加
'                    colDocIDs.Add resultDict
'                End If
'            End If
'        End If
'    Next rowB
'
'    ' 結果を確認
'    Dim docEntry As Variant
'    For Each docEntry In colDocIDs
'        Debug.Print "docID: " & docEntry("docID") & ", docrevChange: " & docEntry("docrevChange")
'    Next docEntry
'End Sub

