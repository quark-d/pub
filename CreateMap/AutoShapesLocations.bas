Option Explicit

' 指定されたシートにオートシェイプを作成する関数
Private Sub CreateShapeIfNotExists(ws As Worksheet, shapeName As String, locationKey As String, offsetX As Single, offsetY As Single, shapeSize As Single)
    Dim existingShape As shape
    Dim cell As Range
    Dim xPos As Single
    Dim yPos As Single
    Dim shape As shape
    
    ' 同じ名前のオートシェイプが存在するか確認
    On Error Resume Next
    Set existingShape = ws.Shapes(shapeName)
    On Error GoTo 0
    
    ' 同じ名前のオートシェイプが存在しない場合、新しいオートシェイプを作成
    If existingShape Is Nothing Then
        ' 仮の位置を設定（適切なセルや位置に設定する必要があります）
        Set cell = ws.Cells(1, 1)
        xPos = cell.Left + offsetX
        yPos = cell.Top + offsetY
        
        ' オートシェイプの作成
        Set shape = ws.Shapes.AddShape(msoShapeOval, xPos, yPos, shapeSize, shapeSize)
        shape.Name = shapeName
        shape.Fill.ForeColor.RGB = RGB(255, 0, 0) ' 赤色
        shape.Line.Visible = msoFalse ' 線なし
        shape.TextFrame2.TextRange.Text = locationKey ' オートシェイプにテキストを追加
    End If
End Sub

' 各シートにオートシェイプを作成する関数
Public Sub AddShapesToSheets(classroomDict As Object, offsetX As Single, offsetY As Single, shapeSize As Single)
    Dim ws As Worksheet
    Dim classKey As Variant
    Dim locationDict As Object
    Dim locationKey As Variant
    
    ' 教室ごとのシートにオートシェイプを追加
    For Each classKey In classroomDict.Keys
        ' classKey を String 型として扱う
        Dim classKeyStr As String
        classKeyStr = CStr(classKey)
        
        ' シートの参照
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(classKeyStr)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' 教室内の場所辞書を取得
            Set locationDict = classroomDict(classKey)
            
            ' 場所ごとにオートシェイプを作成
            For Each locationKey In locationDict.Keys
                ' locationKey を String 型として扱う
                Dim locationKeyStr As String
                locationKeyStr = CStr(locationKey)
                
                ' オートシェイプの名前を決定（クラス名と場所を含める）
                Dim shapeName As String
                shapeName = "shp_" & classKeyStr & "_" & locationKeyStr
                
                ' オートシェイプの作成（存在しない場合のみ）
                Call CreateShapeIfNotExists(ws, shapeName, locationKeyStr, offsetX, offsetY, shapeSize)
            Next locationKey
        End If
    Next classKey
End Sub

' UpdateShapesVisibility関数
Public Sub UpdateShapesVisibility(ByVal boardID As String, ByVal idDict As Object)
    Dim ws As Worksheet
    Dim locationDict As Object
    Dim locationKey As Variant
    Dim locationList As Object
    Dim location As Variant
    Dim locationFullPath As String

    ' 指定されたboardIDが辞書に存在するか確認
    If Not idDict.Exists(boardID) Then
        MsgBox "指定されたboardIDは辞書に存在しません。", vbExclamation
        Exit Sub
    End If

    ' すべてのシートのオートシェイプを非表示にする
    Call SetAllSheetsShapesVisibility(False)
    
    ' 該当する場所の辞書を取得
    Set locationDict = idDict(boardID)

    ' locationDict内の各場所をループ
    For Each locationKey In locationDict.Keys
        For Each location In locationDict(locationKey)
            locationFullPath = "shp_" & locationKey & "_" & location
            ' 該当するオートシェイプをすべてのシートで表示する関数を呼び出す
            Call ShowShapeInAllSheets(locationFullPath)
        Next location
    Next locationKey
End Sub

' すべてのシートで指定したオートシェイプを表示する関数
Private Sub ShowShapeInAllSheets(ByVal shapeName As String)
    Dim ws As Worksheet
    Dim shp As shape
    
    ' すべてのシートをループ
    For Each ws In ThisWorkbook.Sheets
        ' シート内のすべてのオートシェイプをループ
        For Each shp In ws.Shapes
            ' オートシェイプの名前が指定した名前と一致する場合
            If shp.Name = shapeName Then
                shp.Visible = msoTrue
            End If
        Next shp
    Next ws
End Sub


Private Sub SetShapesVisibility(ByVal ws As Worksheet, ByVal visibility As Boolean)
    Dim shp As shape
    
    For Each shp In ws.Shapes
        If shp.Name Like "shp_" & "*" Then
            shp.Visible = IIf(visibility, msoTrue, msoFalse)
        End If
    Next shp
End Sub

Public Sub SetAllSheetsShapesVisibility(ByVal visibility As Boolean)
    Dim ws As Worksheet
    
    ' すべてのシートをループ
    For Each ws In ThisWorkbook.Sheets
        ' 各シートのオートシェイプの表示/非表示を設定
        Call SetShapesVisibility(ws, visibility)
    Next ws
End Sub


