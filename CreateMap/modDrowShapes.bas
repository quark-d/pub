Attribute VB_Name = "modDrowShapes"
Option Explicit

Private Const DELIMITER_ADDR As String = "/"
Private Const DELIMITER_ADDRS As String = ":"



'Sub DrowPosition(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = "/")
'    Dim ws As Worksheet
'    Dim shapeName As String
'    Dim sheetName As String
'    Dim position As String
'    Dim shapeID As String
'    Dim shape As shape
'    Dim xPos As Double
'    Dim yPos As Double
'    Dim redValue As Integer
'    Dim greenValue As Integer
'    Dim blueValue As Integer
'    Dim shapeDict As Dictionary
'    Dim existingShapes As Collection
'    Dim shp As shape
'
'    ' 位置情報を分解
'    sheetName = Split(positionString, delimiter)(0)
'    position = Split(positionString, delimiter)(1)
'    shapeID = "pos_" & sheetName & "_" & position
'
'    ' 既存の "pos_" プレフィックスのオートシェイプリストを取得
'    Set existingShapes = GetAutoShapeList("pos_")
'
'    ' ConfigPositions から座標と色の情報を取得
'    configPosition.AllocateShapeAttribute shapeID, xPos, yPos, redValue, greenValue, blueValue
'
'    ' すでに描画されているかをチェック
'    For Each shp In existingShapes
'        If shp.Name = shapeID Then
'            ' 背景色のみ変更
'            shp.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
'            Exit Sub ' 処理を終了
'        End If
'    Next shp
'
'    ' 該当シートの確認または作成
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets(sheetName)
'    If ws Is Nothing Then
'        Set ws = ThisWorkbook.Sheets.Add
'        ws.Name = sheetName
'    End If
'    On Error GoTo 0
'
'    ' ConfigShapes の設定を取得
'    Set shapeDict = configShape.GetConfigShapePosition()
'
'    ' オートシェイプの描画
'    Set shape = ws.Shapes.AddShape(shapeDict("SHAPE_POSITION_TYPE"), xPos, yPos, shapeDict("SHAPE_POSITION_WIDTH"), shapeDict("SHAPE_POSITION_HEIGHT"))
'    shape.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
'    shape.Name = shapeID ' シェイプ名を設定
'    shape.TextFrame2.TextRange.Text = position
'End Sub
'Public Sub DrowShapes(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = "/")
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim positionList As Collection
'    Dim row As ListRow
'    Dim positions() As String
'    Dim individualPosition As Variant
'    Dim positionParts() As String
'    Dim i As Integer
'
'    Call DrowStandby(configShape, configPosition, positionList)
'
'    ' シートとテーブルの設定
'    Set ws = ThisWorkbook.Sheets("map")
'    Set tbl = ws.ListObjects("tbl_maps")
'
'    ' テーブルの各行をループ
'    For Each row In tbl.ListRows
'        ' addr 列から位置情報を取得
'        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, ":")
'
'        ' それぞれの位置をループして DrowPosition を呼び出す
'        For i = LBound(positions) To UBound(positions)
'            individualPosition = positions(i)
'            ' DrowPosition を呼び出す
'            Call DrowShape(configShape, configPosition, CStr(individualPosition))
'        Next i
'    Next row
'
'End Sub

Public Sub DrowShapes(configShape As ConfigShapes, configPosition As ConfigPositions, positionList As Collection, Optional delimiter As String = DELIMITER_ADDR)
    Dim individualPosition As Variant
    
    ' 準備関数の呼び出し
    Call DrowStandby(configShape, configPosition, positionList)
    
    ' 位置リストをループして DrowShape を呼び出す
    For Each individualPosition In positionList
        ' DrowShape を呼び出す
        Call DrowShape(configShape, configPosition, CStr(individualPosition), delimiter)
    Next individualPosition
End Sub


'Private Sub DrowShape(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = "/")
'    Dim ws As Worksheet
'    Dim labelName As String
'    Dim sheetName As String
'    Dim position As String
'    Dim shapeID As String
'    Dim xPos As Double
'    Dim yPos As Double
'    Dim redValue As Integer
'    Dim greenValue As Integer
'    Dim blueValue As Integer
'    Dim shapeDict As Dictionary
'    Dim existingShapes As Collection
'    Dim existingShape As shape
'    Dim shp As shape
'
'    ' 位置情報を分解
'    sheetName = Split(positionString, delimiter)(0)
'    position = Split(positionString, delimiter)(1)
'    shapeID = "pos_" & sheetName & "_" & position
'    labelName = shapeID & "_lbl"
'
'    ' 既存の "pos_" プレフィックスのオートシェイプリストを取得
'    Set existingShapes = GetAutoShapeListPosition()
'
'    ' ConfigPositions から座標と色の情報を取得
'    configPosition.AllocateShapeAttribute shapeID, xPos, yPos, redValue, greenValue, blueValue
'
'
'    ' 該当シートの確認または作成
'    On Error Resume Next
'    Set ws = ThisWorkbook.Sheets(sheetName)
'    If ws Is Nothing Then
'        Set ws = ThisWorkbook.Sheets.Add
'        ws.Name = sheetName
'    End If
'    On Error GoTo 0
'
'    ' ConfigShapes の設定を取得
'    Set shapeDict = configShape.GetConfigShapePosition()
'
'
'    ' すでに描画されているかをチェック
'    Set existingShape = Nothing
'    For Each shp In existingShapes
'        If shp.Name = shapeID Then
'            Set existingShape = shp
'            ' 背景色のみ変更
'            shp.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
'
'            ' ラベルとコネクタを描画
'            On Error Resume Next ' エラー処理を追加
'            Call DrowLabel(ws, shapeDict, labelName, shp.left, shp.top)
'            Call DrowConnection(ws, shp, ws.Shapes(labelName))
'            On Error GoTo 0
'
'            Exit Sub ' 処理を終了
'        End If
'    Next shp
'
'          ' positionとlabelを描画し、直線で結ぶ
'          Call DrowPosition(ws, shapeDict, shapeID, xPos, yPos, redValue, greenValue, blueValue, position)
'          Call DrowLabel(ws, shapeDict, labelName, xPos, yPos)
'          Call DrowConnection(ws, ws.Shapes(shapeID), ws.Shapes(labelName))
'
'End Sub
Private Sub DrowShape(configShape As ConfigShapes, configPosition As ConfigPositions, positionString As Variant, Optional delimiter As String = DELIMITER_ADDR)
    Dim ws As Worksheet
    Dim labelName As String
    Dim sheetName As String
    Dim position As String
    Dim shapeID As String
    Dim xPos As Double
    Dim yPos As Double
    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer
    Dim shapeDict As Dictionary
    Dim existingShapes As Collection
    Dim existingShape As shape
    Dim shp As shape
    Dim labelShapes As Collection
    Dim existingLabel As shape

    ' 位置情報を分解
    sheetName = Split(positionString, delimiter)(0)
    position = Split(positionString, delimiter)(1)
    shapeID = "pos_" & sheetName & "_" & position
    labelName = shapeID & "_lbl"

    ' 既存の "pos_" プレフィックスのオートシェイプリストを取得
    Set existingShapes = GetAutoShapeListPosition()
    
    ' ConfigPositions から座標と色の情報を取得
    configPosition.AllocateShapeAttribute shapeID, xPos, yPos, redValue, greenValue, blueValue

    ' 該当シートの確認または作成
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    End If
    On Error GoTo 0

    ' ConfigShapes の設定を取得
    Set shapeDict = configShape.GetConfigShapePosition()
    
    ' すでに描画されているかをチェック
    Set existingShape = Nothing
    For Each shp In existingShapes
        If shp.Name = shapeID Then
            Set existingShape = shp
            ' 背景色のみ変更
            shp.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
            
            ' ラベルシェイプの存在を確認
            Set labelShapes = GetAutoShapeListLabel()
            Set existingLabel = Nothing
            For Each existingLabel In labelShapes
                If existingLabel.Name = labelName Then
                    Exit Sub ' ラベルが存在するため処理を終了
                End If
            Next existingLabel
            
            ' ラベルとコネクタを描画
            Call DrowLabel(ws, shapeDict, labelName, shp.left, shp.top)
            Call DrowConnection(ws, shp, ws.Shapes(labelName))
            
            Exit Sub ' 処理を終了
        End If
    Next shp
          
    ' positionとlabelを描画し、直線で結ぶ
    Call DrowPosition(ws, shapeDict, shapeID, xPos, yPos, redValue, greenValue, blueValue, position)
    Call DrowLabel(ws, shapeDict, labelName, xPos, yPos)
    Call DrowConnection(ws, ws.Shapes(shapeID), ws.Shapes(labelName))

End Sub



Private Sub DrowPosition(ws As Worksheet, shapeDict As Dictionary, shapeID As String, xPos As Double, yPos As Double, redValue As Integer, greenValue As Integer, blueValue As Integer, position As String)
    Dim shape As shape
    ' positionオートシェイプの描画
    Set shape = ws.Shapes.AddShape(shapeDict("SHAPE_POSITION_TYPE"), xPos, yPos, shapeDict("SHAPE_POSITION_WIDTH"), shapeDict("SHAPE_POSITION_HEIGHT"))
    shape.Fill.ForeColor.RGB = RGB(redValue, greenValue, blueValue)
    shape.Name = shapeID ' シェイプ名を設定
    shape.TextFrame2.TextRange.Text = position
End Sub
Private Sub DrowLabel(ws As Worksheet, shapeDict As Dictionary, labelName As String, xPos As Double, yPos As Double)
    Dim label As shape
    Dim labelType As MsoAutoShapeType
    'Dim labelType As String
    Dim left As Single
    Dim top As Single
    Dim width As Single
    Dim height As Single

    'labelType = shapeDict("SHAPE_LABEL_TYPE")
    left = xPos
    top = yPos + shapeDict("SHAPE_POSITION_HEIGHT") + 10
    width = shapeDict("SHAPE_LABEL_WIDTH")
    height = shapeDict("SHAPE_LABEL_HEIGHT")
    
    ' labelオートシェイプの描画（positionの下に表示）
    Set label = ws.Shapes.AddShape(shapeDict("SHAPE_LABEL_TYPE"), left, top, width, height)
    label.Fill.ForeColor.RGB = RGB(200, 200, 200) ' 初期は薄い灰色
    label.Name = labelName
    label.TextFrame2.TextRange.Text = "" ' 空文字
    label.Line.Visible = msoFalse ' 外周の線を非表示
End Sub

Private Sub DrowConnection(ws As Worksheet, shape As shape, label As shape)
    Dim connector As shape
    ' positionとlabelを直線で結ぶ
    Set connector = ws.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
    connector.ConnectorFormat.BeginConnect shape, 2 ' positionシェイプの中央下
    connector.ConnectorFormat.EndConnect label, 1   ' labelシェイプの中央上
    connector.Line.Weight = 1.5 ' 線の太さ
End Sub

Private Function GetAutoShapeListPosition() As Collection
    Dim shapeList As New Collection
    Dim ws As Worksheet
    Dim shp As shape
    
    ' 全シートをループ
    For Each ws In ThisWorkbook.Worksheets
        ' シート内のすべてのシェイプをループ
        For Each shp In ws.Shapes
            ' "pos_" プレフィックスで始まり、"_lbl" サフィックスで終わらない場合、リストに追加
            If shp.Name Like "pos_*" And Not shp.Name Like "*_lbl" Then
                shapeList.Add shp
            End If
        Next shp
    Next ws
    
    ' 結果を返す
    Set GetAutoShapeListPosition = shapeList
End Function

Public Function GetAutoShapeListLabel() As Collection
    Dim shapeList As New Collection
    Dim ws As Worksheet
    Dim shp As shape
    
    ' 全シートをループ
    For Each ws In ThisWorkbook.Worksheets
        ' シート内のすべてのシェイプをループ
        For Each shp In ws.Shapes
            ' "pos_" プレフィックスで始まり、"_lbl" サフィックスで終わる場合、リストに追加
            If shp.Name Like "pos_*_lbl" Then
                shapeList.Add shp
            End If
        Next shp
    Next ws
    
    ' 結果を返す
    Set GetAutoShapeListLabel = shapeList
End Function


Private Sub DrowStandby(configShape As ConfigShapes, configPosition As ConfigPositions, positionList As Collection)
  Dim existingShapes As Collection
  Dim shp As shape
  Dim labelName As String
  
  ' "pos_" から始まるすべてのシェイプを取得
  Set existingShapes = GetAutoShapeListPosition()
  
  ' すべての "pos_" から始まるシェイプをループ
  For Each shp In existingShapes
      ' ラベル用のサフィックスを持つシェイプ名を作成
      labelName = shp.Name & "_lbl"
  
      ' ラベルオートシェイプを削除
      On Error Resume Next
      shp.Parent.Shapes(labelName).Delete
      On Error GoTo 0
  
      ' 接続シェイプを削除
'      Dim connector As shape
'      For Each connector In shp.Parent.Shapes
'          If connector.Type = MsoConnector Then
'              If connector.ConnectorFormat.BeginConnectedShape Is shp Or connector.ConnectorFormat.EndConnectedShape Is shp Then
'                  connector.Delete
'              End If
'          End If
'      Next connector
'      Dim connector As shape
'      For Each connector In shp.Parent.Shapes
'          If connector.connector Then
'              ' シェイプがコネクタであり、現在のシェイプに接続されているかをチェック
'              If connector.ConnectorFormat.BeginConnectedShape Is shp Or connector.ConnectorFormat.EndConnectedShape Is shp Then
'                  connector.Delete
'              End If
'          End If
'      Next connector
        Dim connector As shape
        For Each connector In shp.Parent.Shapes
            ' シェイプがコネクタかどうかを確認
            If connector.connector Then
                ' コネクタが現在のシェイプに接続されているかをチェック
                If Not connector.ConnectorFormat.BeginConnectedShape Is Nothing Then
                    If connector.ConnectorFormat.BeginConnectedShape Is shp Then
                        connector.Delete
                    End If
'                    If connector.ConnectorFormat.EndConnectedShape Is shp Then
'                      connector.Delete
'                    End If
                End If
            End If
        Next connector
  
      ' 残ったシェイプを薄い灰色に設定
      shp.Fill.ForeColor.RGB = RGB(200, 200, 200)
  Next shp

End Sub
Sub AppendTextInLabel(tbl As ListObject, configShape As ConfigShapes, selectedIDs As Collection)
    Dim ws As Worksheet
    Dim row As ListRow
    Dim positions() As String
    Dim individualPosition As Variant
    Dim sheetName As String
    Dim position As String
    Dim shapeID As String
    Dim labelName As String
    Dim labelShape As shape
    Dim posShape As shape
    Dim shapeDict As Dictionary
    Dim xPos As Double
    Dim yPos As Double
    Dim labelShapes As Collection
    Dim shp As shape
    
    ' ConfigShapes の設定を取得
    Set shapeDict = configShape.GetConfigShapePosition()
    
    ' テーブルの各行をループ
    For Each row In tbl.ListRows
        ' addr 列から位置情報を取得
        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, DELIMITER_ADDRS)
        
        ' それぞれの位置をループ
        For Each individualPosition In positions
            sheetName = Split(individualPosition, DELIMITER_ADDR)(0)
            position = Split(individualPosition, DELIMITER_ADDR)(1)
            shapeID = "pos_" & sheetName & "_" & position
            labelName = shapeID & "_lbl"

            ' シートの取得または作成
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(sheetName)
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Sheets.Add
                ws.Name = sheetName
            End If
            On Error GoTo 0

            ' 既存のラベルシェイプリストを取得
            Set labelShapes = GetAutoShapeListLabel()
            
            ' 既存のラベルシェイプを探す
            Set labelShape = Nothing
            For Each shp In labelShapes
                If shp.Name = labelName Then
                    Set labelShape = shp
                    Exit For
                End If
            Next shp

            ' ラベルシェイプが見つからない場合、新規作成をスキップ
            If labelShape Is Nothing Then
                ' 次の位置情報にスキップ
                GoTo ContinueLoop
            End If

            ' selectedIDs と一致するか確認
            Dim selectedID As Variant
            For Each selectedID In selectedIDs
                If row.Range(1).value = selectedID Then
                    ' ラベルシェイプにテキストを追加
                    Call SetSheetTabColor(labelShape.Parent)
                    If Not labelShape Is Nothing Then
                        If labelShape.TextFrame2.TextRange.Text <> "" Then
                            labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & vbCrLf
                        End If
                        labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & row.Range(1).value
                    End If
                    Exit For
                End If
            Next selectedID

ContinueLoop:
        Next individualPosition
    Next row
End Sub

'Sub AppendTextInLabel(tbl As ListObject, configShape As ConfigShapes)
'    Dim ws As Worksheet
'    Dim row As ListRow
'    Dim positions() As String
'    Dim individualPosition As Variant
'    Dim sheetName As String
'    Dim position As String
'    Dim shapeID As String
'    Dim labelName As String
'    Dim labelShape As shape
'    Dim posShape As shape
'    Dim shapeDict As Dictionary
'    Dim xPos As Double
'    Dim yPos As Double
'    Dim labelShapes As Collection
'    Dim shp As shape
'
'    ' ConfigShapes の設定を取得
'    Set shapeDict = configShape.GetConfigShapePosition()
'
'    ' テーブルの各行をループ
'    For Each row In tbl.ListRows
'        ' addr 列から位置情報を取得
'        positions = Split(row.Range(tbl.ListColumns("addr").Index).value, ":")
'
'        ' それぞれの位置をループ
'        For Each individualPosition In positions
'            sheetName = Split(individualPosition, "/")(0)
'            position = Split(individualPosition, "/")(1)
'            shapeID = "pos_" & sheetName & "_" & position
'            labelName = shapeID & "_lbl"
'
'            ' シートの取得または作成
'            On Error Resume Next
'            Set ws = ThisWorkbook.Sheets(sheetName)
'            If ws Is Nothing Then
'                Set ws = ThisWorkbook.Sheets.Add
'                ws.Name = sheetName
'            End If
'            On Error GoTo 0
'
'            ' 既存のラベルシェイプリストを取得
'            Set labelShapes = GetAutoShapeListLabel()
'
'            ' 既存のラベルシェイプを探す
'            Set labelShape = Nothing
'            For Each shp In labelShapes
'                If shp.Name = labelName Then
'                    Set labelShape = shp
'                    Exit For
'                End If
'            Next shp
'
'            ' ラベルシェイプが見つからない場合、新規作成
'            If labelShape Is Nothing Then
'                ' Positionシェイプの位置を取得
'                On Error Resume Next
'                Set posShape = ws.Shapes(shapeID)
'                On Error GoTo 0
'
'                If Not posShape Is Nothing Then
'                    ' Positionシェイプの下にラベルを作成
'                    xPos = posShape.left
'                    yPos = posShape.top + posShape.height + 10
'                    Set labelShape = ws.Shapes.AddShape(shapeDict("SHAPE_LABEL_TYPE"), xPos, yPos, shapeDict("SHAPE_LABEL_WIDTH"), shapeDict("SHAPE_LABEL_HEIGHT"))
'                    labelShape.Name = labelName
'                    labelShape.Fill.ForeColor.RGB = RGB(200, 200, 200)
'                    labelShape.TextFrame2.TextRange.Text = ""
'                    labelShape.Line.Visible = msoFalse ' 外周の線を非表示
'                End If
'            End If
'
'            ' ラベルシェイプにテキストを追加
'            If Not labelShape Is Nothing Then
'                If labelShape.TextFrame2.TextRange.Text <> "" Then
'                    labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & vbCrLf
'                End If
'                labelShape.TextFrame2.TextRange.Text = labelShape.TextFrame2.TextRange.Text & row.Range(1).value
'            End If
'        Next individualPosition
'    Next row
'End Sub
'
'Public Sub AdjustShapeLabelByText(labelShape As shape)
'    Dim textLines() As String
'    Dim maxWidth As Double
'    Dim i As Integer
'
'    ' 複数行のテキストを考慮して最も長い行の幅を計算
'    textLines = Split(labelShape.TextFrame2.TextRange.Text, vbCrLf)
'    maxWidth = 0
'    For i = LBound(textLines) To UBound(textLines)
'        labelShape.TextFrame2.TextRange.Text = textLines(i)
'        If labelShape.TextFrame2.TextRange.BoundWidth > maxWidth Then
'            maxWidth = labelShape.TextFrame2.TextRange.BoundWidth
'        End If
'    Next i
'
'    ' テキストの高さと最も長い行に合わせてラベルシェイプのサイズを調整
'    labelShape.height = labelShape.TextFrame2.TextRange.BoundHeight + 10 ' 高さを調整
'    labelShape.width = maxWidth + 10   ' 最も長い行に基づいて幅を調整
'End Sub
Sub AdjustShapeLabelByText(labelShape As shape)
    Dim textLines() As String
    Dim maxWidth As Double
    Dim i As Integer
    
    ' 自動折り返しを無効化
    labelShape.TextFrame2.WordWrap = msoFalse
    
    ' 複数行のテキストを考慮して最も長い行の幅を計算
    textLines = Split(labelShape.TextFrame2.TextRange.Text, vbCrLf)
    maxWidth = 0
    For i = LBound(textLines) To UBound(textLines)
        labelShape.TextFrame2.TextRange.Text = textLines(i)
        If labelShape.TextFrame2.TextRange.BoundWidth > maxWidth Then
            maxWidth = labelShape.TextFrame2.TextRange.BoundWidth
        End If
    Next i
    
    ' テキストの高さと最も長い行に合わせてラベルシェイプのサイズを調整
    labelShape.height = labelShape.TextFrame2.TextRange.BoundHeight + 10 ' 高さを調整
    'labelShape.width = maxWidth + 10   ' 最も長い行に基づいて幅を調整
    
    ' 自動折り返しを再度有効化（必要に応じて）
    labelShape.TextFrame2.WordWrap = msoTrue
End Sub

