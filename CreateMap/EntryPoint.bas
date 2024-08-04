Option Explicit

Sub ExampleUsage()
    Call SetDelimiterAndColumns("/", ":", "ID", "addr")
    Call loadAddressLists(ThisWorkbook, "map", "tbl_maps")
    
    Dim dict As Object
    Dim idDict As Object



    ' 教室ごとの辞書オブジェクトを取得
    Set dict = getClassroomDict()
    ' 掲示板IDごとの辞書オブジェクトを取得
    Set idDict = getIdDict()
    
    Call DumpParClasss(dict)
    Call DumpParID(idDict)
    
    ' シートの確認
    Call CheckAndAddSheets
    ' Call CheckAndAddSheetsById
    
    ' マーカーを付ける
    Call RunAddShapesToSheets
    
    Call UpdateShapesVisibility("jp-002", idDict)
    
    
    Dim result As Collection
    Set result = GetBoardIDsForLocationFromClassroomDict("shp_classB_1F_abc")
    
End Sub




Sub RunAddShapesToSheets()
    Dim classroomDict As Object
    Dim offsetX As Single
    Dim offsetY As Single
    Dim shapeSize As Single
    
    ' classroomDict を getClassroomDict() などで取得する処理
    Set classroomDict = getClassroomDict()
    
    ' オートシェイプのサイズと位置のオフセット
    offsetX = 5
    offsetY = 5
    shapeSize = 40
    
    ' AddShapesToSheets 関数を呼び出し
    Call AddShapesToSheets(classroomDict, offsetX, offsetY, shapeSize)
End Sub

Sub ShowAllShapes()
    Call SetAllSheetsShapesVisibility(True)
End Sub


Public Sub ShowSelectFormRadio()
    Call SetDelimiterAndColumns("/", ":", "ID", "addr")
    Call loadAddressLists(ThisWorkbook, "map", "tbl_maps")
    
    Dim dict As Object
    Dim idDict As Object

    ' 教室ごとの辞書オブジェクトを取得
    Set dict = getClassroomDict()
    ' 掲示板IDごとの辞書オブジェクトを取得
    Set idDict = getIdDict()
    
    Call DumpParClasss(dict)
    Call DumpParID(idDict)
    
    ' シートの確認
    Call CheckAndAddSheets
    ' Call CheckAndAddSheetsById
    
    ' マーカーを付ける
    Call RunAddShapesToSheets
    
    frmSelectIDRadio.Show
    
    
End Sub

Public Sub ShowSelectFormList()
    Call SetDelimiterAndColumns("/", ":", "ID", "addr")
    Call loadAddressLists(ThisWorkbook, "map", "tbl_maps")
    
    Dim dict As Object
    Dim idDict As Object

    ' 教室ごとの辞書オブジェクトを取得
    Set dict = getClassroomDict()
    ' 掲示板IDごとの辞書オブジェクトを取得
    Set idDict = getIdDict()
    
    Call DumpParClasss(dict)
    Call DumpParID(idDict)
    
    ' シートの確認
    Call CheckAndAddSheets
    ' Call CheckAndAddSheetsById
    
    ' マーカーを付ける
    Call RunAddShapesToSheets
    
    frmSelectIDList.Show
    
    
End Sub
