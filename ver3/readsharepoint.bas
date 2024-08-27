Function RetrieveRecordsFromSharePoint() As Collection
    ' 変数宣言
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim strConn As String
    Dim sharePointURL As String
    Dim listName As String
    Dim colRecords As New Collection
    Dim recordDict As Object

    ' configConnectシートとtbl_config_conテーブルの確認
    On Error Resume Next
    Dim wsConfig As Worksheet
    Dim tblConfig As ListObject
    Set wsConfig = ThisWorkbook.Sheets("configConnect")
    On Error GoTo 0

    ' シートが存在しない場合はエラーを返す
    If wsConfig Is Nothing Then
        MsgBox "configConnectシートが見つかりません。", vbCritical
        Exit Function
    End If

    ' テーブルの確認
    On Error Resume Next
    Set tblConfig = wsConfig.ListObjects("tbl_config_con")
    On Error GoTo 0

    ' テーブルが存在しない場合はエラーを返す
    If tblConfig Is Nothing Then
        MsgBox "tbl_config_conテーブルが見つかりません。", vbCritical
        Exit Function
    End If

    ' URLとリスト名の取得
    sharePointURL = tblConfig.DataBodyRange.Cells(1, 1).Value
    listName = tblConfig.DataBodyRange.Cells(1, 2).Value

    ' URLまたはリスト名が空の場合はメッセージを表示して終了
    If sharePointURL = "" Or listName = "" Then
        MsgBox "URLまたはリスト名が指定されていません。", vbCritical
        Exit Function
    End If

    ' 接続文字列の作成
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "WSS;IMEX=1;RetrieveIds=Yes;" & _
              "DATABASE=" & sharePointURL

    ' 接続オブジェクトの作成
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn

    ' RecordSetオブジェクトの作成
    Set rs = CreateObject("ADODB.Recordset")

    ' SharePointリストのデータを取得
    rs.Open "SELECT ID_ID, COL1, COL2 FROM [" & listName & "]", conn, 1, 3

    ' レコードをループしてCollectionに追加
    Do Until rs.EOF
        Set recordDict = CreateObject("Scripting.Dictionary")
        recordDict("ID_ID") = rs.Fields("ID_ID").Value
        recordDict("COL1") = rs.Fields("COL1").Value
        recordDict("COL2") = rs.Fields("COL2").Value
        colRecords.Add recordDict
        rs.MoveNext
    Loop

    ' RecordSetとConnectionを閉じる
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' コレクションを返す
    Set RetrieveRecordsFromSharePoint = colRecords
End Function
