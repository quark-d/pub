Option Explicit

' シート名が存在するか確認し、存在しない場合は追加する関数
Public Sub CheckAndAddSheets()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim classroomDict As Object
    Dim key As Variant
    
    ' 教室ごとの辞書オブジェクトを取得
    Set classroomDict = getClassroomDict()
    
    ' 教室ごとのシートをチェックし、存在しない場合は追加
    For Each key In classroomDict.Keys
        sheetExists = False
        
        ' シートの存在確認
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = key Then
                sheetExists = True
                Exit For
            End If
        Next ws
        
        ' シートが存在しない場合、新しいシートを追加
        If Not sheetExists Then
            ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = key
        End If
    Next key
End Sub

' IDごとの辞書オブジェクトからシート名をチェックして追加する関数
Public Sub CheckAndAddSheetsById()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim idDict As Object
    Dim key As Variant
    Dim classroomDict As Object
    Dim classKey As Variant
    
    ' 掲示板IDごとの辞書オブジェクトを取得
    Set idDict = getIdDict()
    
    ' 掲示板IDごとのシートをチェックし、存在しない場合は追加
    For Each key In idDict.Keys
        ' 掲示板IDに関連する教室のチェック
        For Each classKey In idDict(key).Keys
            sheetExists = False
            
            ' シートの存在確認
            For Each ws In ThisWorkbook.Sheets
                If ws.Name = classKey Then
                    sheetExists = True
                    Exit For
                End If
            Next ws
            
            ' シートが存在しない場合、新しいシートを追加
            If Not sheetExists Then
                ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = classKey
            End If
        Next classKey
    Next key
End Sub

