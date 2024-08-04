Option Explicit


' モジュール内のプライベート変数として区切り文字と辞書オブジェクトを定義
Private classroomDelimiter As String
Private placeDelimiter As String
Private idColumn As String
Private addrColumn As String
Private classroomDict As Object ' 教室ごとの辞書オブジェクト
Private idDict As Object ' 掲示板IDごとの辞書オブジェクト

' 区切り文字と列名を設定する関数
Public Sub SetDelimiterAndColumns(ByVal clsDelimiter As String, ByVal plDelimiter As String, ByVal idCol As String, ByVal addrCol As String)
    classroomDelimiter = clsDelimiter
    placeDelimiter = plDelimiter
    idColumn = idCol
    addrColumn = addrCol
End Sub

' 列番号を取得するヘルパー関数
Private Sub GetColumnIndexes(ByVal tbl As ListObject, ByRef idColumnIndex As Integer, ByRef addrColumnIndex As Integer)
    idColumnIndex = tbl.ListColumns(idColumn).Index
    addrColumnIndex = tbl.ListColumns(addrColumn).Index
End Sub

' 共通処理を含む関数
Private Sub ProcessTable(ByVal Workbook As Workbook, ByVal sheetName As String, ByVal tableName As String)
    Dim tbl As ListObject
    Dim r As ListRow
    Dim addrList() As String
    Dim i As Integer
    Dim classroom As String
    Dim location As String
    Dim idColumnIndex As Integer
    Dim addrColumnIndex As Integer
    
    ' テーブルの参照
    Set tbl = Workbook.Sheets(sheetName).ListObjects(tableName)
    
    ' 列番号を取得
    Call GetColumnIndexes(tbl, idColumnIndex, addrColumnIndex)
    
    ' テーブルの各行を処理
    For Each r In tbl.ListRows
        ' addr列を取得し、場所の区切り文字で分割
        addrList = Split(r.Range.Cells(1, addrColumnIndex).Value, placeDelimiter)
        
        ' CreateClassroomDictとCreateIdDictでの処理を行う
        Call CreateClassroomDict(r, addrList, idColumnIndex, classroomDelimiter, placeDelimiter)
        Call CreateIdDict(r, addrList, idColumnIndex, classroomDelimiter, placeDelimiter)
    Next r
End Sub

' 教室ごとのリストを作成する関数
Private Sub CreateClassroomDict(ByRef r As ListRow, ByRef addrList() As String, ByVal idColumnIndex As Integer, ByVal classroomDelimiter As String, ByVal placeDelimiter As String)
    Dim places() As String
    Dim i As Integer
    Dim classroom As String
    Dim location As String
    Dim key As String
    Dim locationDict As Object
    
    ' 教室ごとの辞書オブジェクトの作成
    If classroomDict Is Nothing Then
        Set classroomDict = CreateObject("Scripting.Dictionary")
    End If
    
    For i = LBound(addrList) To UBound(addrList)
        ' 教室名を分割し、最初の部分を取得
        classroom = Split(addrList(i), classroomDelimiter)(0)
        
        ' 教室名を辞書のキーとして使用
        If Not classroomDict.Exists(classroom) Then
            Set classroomDict(classroom) = CreateObject("Scripting.Dictionary")
        End If
        
        ' 場所の部分を分割し、辞書に追加
        places = Split(Split(addrList(i), classroomDelimiter)(1), placeDelimiter)
        location = places(0)
        
        ' 教室内の場所辞書を作成
        If Not classroomDict(classroom).Exists(location) Then
            Set locationDict = CreateObject("Scripting.Dictionary")
            locationDict.Add r.Range.Cells(1, idColumnIndex).Value, location
            classroomDict(classroom).Add location, locationDict
        Else
            Set locationDict = classroomDict(classroom)(location)
            If Not locationDict.Exists(r.Range.Cells(1, idColumnIndex).Value) Then
                locationDict.Add r.Range.Cells(1, idColumnIndex).Value, location
            End If
        End If
    Next i
End Sub

' 掲示板IDごとのリストを作成する関数
Private Sub CreateIdDict(ByRef r As ListRow, ByRef addrList() As String, ByVal idColumnIndex As Integer, ByVal classroomDelimiter As String, ByVal placeDelimiter As String)
    Dim places() As String
    Dim i As Integer
    Dim classroom As String
    Dim location As String
    Dim key As String
    Dim classDict As Object
    
    ' 掲示板IDごとの辞書オブジェクトの作成
    If idDict Is Nothing Then
        Set idDict = CreateObject("Scripting.Dictionary")
    End If
    
    ' 掲示板IDごとの辞書オブジェクトを更新
    If Not idDict.Exists(r.Range.Cells(1, idColumnIndex).Value) Then
        Set idDict(r.Range.Cells(1, idColumnIndex).Value) = CreateObject("Scripting.Dictionary")
    End If
    
    For i = LBound(addrList) To UBound(addrList)
        ' 教室名を分割し、最初の部分を取得
        classroom = Split(addrList(i), classroomDelimiter)(0)
        
        ' 教室名を辞書のキーとして使用
        If Not idDict(r.Range.Cells(1, idColumnIndex).Value).Exists(classroom) Then
            Set classDict = CreateObject("Scripting.Dictionary")
            classDict.Add Split(Split(addrList(i), classroomDelimiter)(1), placeDelimiter)(0), r.Range.Cells(1, idColumnIndex).Value
            idDict(r.Range.Cells(1, idColumnIndex).Value).Add classroom, classDict
        Else
            Set classDict = idDict(r.Range.Cells(1, idColumnIndex).Value)(classroom)
            location = Split(Split(addrList(i), classroomDelimiter)(1), placeDelimiter)(0)
            If Not classDict.Exists(location) Then
                classDict.Add location, r.Range.Cells(1, idColumnIndex).Value
            End If
        End If
    Next i
End Sub

' loadAddressLists関数で共通処理を行い、その後CreateClassroomDictとCreateIdDictを呼び出す
Public Sub loadAddressLists(ByVal Workbook As Workbook, ByVal sheetName As String, ByVal tableName As String)
    ' 共通処理を含む関数を呼び出し
    Call ProcessTable(Workbook, sheetName, tableName)
End Sub

' 教室ごとの辞書オブジェクトを取得する関数
Public Function getClassroomDict() As Object
    Set getClassroomDict = classroomDict
End Function

' 掲示板IDごとの辞書オブジェクトを取得する関数
Public Function getIdDict() As Object
    Set getIdDict = idDict
End Function



' 場所の文字列を引数に、該当する掲示板IDのリストを返却する関数
Public Function GetBoardIDsForLocationFromClassroomDict(ByVal shp_location As String) As Collection
    ' Dim boardIDs As Object
    Dim boardIDs As Collection
    Dim classroomDict As Object
    Dim classroomKey As Variant
    Dim locationKey As Variant
    Dim locationDict As Object
    Dim location As Variant
    Dim locationFullPath As String
    Dim idDict As Object
    Dim boardID As Variant
    Dim resultList() As String
    Dim i As Integer
    
    ' 結果を格納するコレクションを作成
    ' Set boardIDs = CreateObject("Scripting.Dictionary")
    Set boardIDs = New Collection
    
    ' classroomDictを取得
    Set classroomDict = getClassroomDict()
    
    ' 各教室をループ
    ' classA_1F
    For Each classroomKey In classroomDict.Keys
        ' 各教室の場所辞書を取得
        Set locationDict = classroomDict(classroomKey)
        ' 各場所をループ
        ' ab
        For Each locationKey In locationDict.Keys
            locationFullPath = "shp_" & classroomKey & "_" & locationKey
            If locationFullPath = shp_location Then
                For Each boardID In locationDict(locationKey)
                    boardIDs.Add boardID
                Next boardID
            End If
        Next locationKey
    Next classroomKey
    Set GetBoardIDsForLocationFromClassroomDict = boardIDs
End Function






Public Sub DumpParClasss(ByVal dict As Object)
    Dim classDict As Object
    Dim locDict As Object
    Dim key As Variant
    Dim locKey As Variant
    Dim item As Variant
    Dim result As String

    ' 教室ごとの辞書オブジェクトの内容を確認
    For Each key In dict.Keys
        Debug.Print "Classroom: " & key
        Set classDict = dict(key)
        For Each locKey In classDict.Keys
            result = ""
            For Each item In classDict(locKey)
                If result = "" Then
                    result = item
                Else
                    result = result & ", " & item
                End If
            Next item
            Debug.Print "  Location: " & locKey & " - " & result
        Next locKey
    Next key

End Sub

Public Sub DumpParID(ByVal dict As Object)
    Dim locDict As Object
    Dim key As Variant
    Dim locKey As Variant
    Dim item As Variant
    Dim result As String
    
    ' 掲示板IDごとの辞書オブジェクトの内容を確認
    For Each key In dict.Keys
        Debug.Print "ID: " & key
        Set locDict = dict(key)
        For Each locKey In locDict.Keys
            result = ""
            For Each item In locDict(locKey)
                If result = "" Then
                    result = item
                Else
                    result = result & ", " & item
                End If
            Next item
            Debug.Print "  Location: " & locKey & " - " & result
        Next locKey
    Next key
End Sub
