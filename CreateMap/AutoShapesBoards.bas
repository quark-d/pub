Option Explicit

Public Sub DrawShapesForDisplay()
    Dim ws As Worksheet
    Dim shp As shape
    Dim classroomDict As Object
    Dim locationKey As Variant
    Dim locationDict As Object
    Dim displayShape As shape
    Dim displayText As String
    Dim topOffset As Single
    Dim shape As shape

    ' クラスルームの辞書を取得
    Set classroomDict = getClassroomDict()

    ' 各シートをループ
    For Each ws In ThisWorkbook.Sheets
        ' 各表示中のオートシェイプをループ
        For Each shp In ws.Shapes
            ' 形状の名前が "shp_" で始まる場合
            If shp.Name Like "shp_*" And shp.Visible = msoTrue Then
                ' 教室名と場所の辞書を取得
                locationKey = GetLocationKeyFromShapeName(shp.Name)
                If classroomDict.Exists(locationKey) Then
                    Set locationDict = classroomDict(locationKey)
                    
                    ' 位置の下に描画する四角形のテキストを作成
                    displayText = Join(locationDict.Keys, vbCrLf)
                    
                    ' 四角形の描画
                    topOffset = shp.Top + shp.Height + 10 ' 現在のオートシェイプの下に配置
                    Set displayShape = ws.Shapes.AddShape(msoShapeRectangle, shp.Left, topOffset, shp.Width, 50)
                    displayShape.Name = "board_" & shp.Name
                    displayShape.TextFrame2.TextRange.Text = displayText
                    displayShape.Line.Visible = msoFalse
                End If
            End If
        Next shp
    Next ws
End Sub

' 教室の辞書から形状名に対応する位置キーを取得する関数
Private Function GetLocationKeyFromShapeName(shapeName As String) As String
    Dim locationKey As String
    ' 形状名から "shp_" を削除し、"_class" 部分までを取得
    locationKey = Split(Split(shapeName, "_")(1), "_")(0) & "_" & Split(Split(shapeName, "_")(1), "_")(1)
    GetLocationKeyFromShapeName = locationKey
End Function

