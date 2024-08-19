Attribute VB_Name = "ExportPDF"
Option Explicit

'Sub RotateAndExportSheetsToPDF(sheetNames As Collection, outputFileName As String)
'    Dim ws As Worksheet
'    Dim shapeArray() As Variant
'    Dim shp As shape
'    Dim shapeGroup As shape
'    Dim i As Integer
'    Dim selectedSheets() As String
'    Dim pdfPath As String
'    Dim sheet As Variant
'    Dim originalRotations As Collection
'    Dim originalOrientations As Collection
'    Dim originalZoom As Collection
'
'    ' 各シートの元の設定を保存するためのコレクションを初期化
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF保存先のパスを指定
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' 引数として受け取ったシート名のコレクションを配列に変換
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' 各シートを処理
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' シート上のすべてのシェイプをグループ化して90度回転
'        If ws.Shapes.Count > 0 Then
'            ReDim shapeArray(1 To ws.Shapes.Count)
'            For i = 1 To ws.Shapes.Count
'                shapeArray(i) = ws.Shapes(i).Name
'            Next i
'
'            Set shapeGroup = ws.Shapes.Range(shapeArray).Group
'
'            ' 元の回転角度を保存
'            originalRotations.Add shapeGroup.Rotation
'
'            ' 90度回転
'            shapeGroup.Rotation = shapeGroup.Rotation + 90
'
'            ' グループがシートの範囲を超えないように再配置
'            If shapeGroup.left < 0 Then shapeGroup.left = 0
'            If shapeGroup.top < 0 Then shapeGroup.top = 0
'        End If
'
'        ' 元のページ設定を保存
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' ページ設定で縮尺を無効にして設定
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' または xlLandscape、必要に応じて設定
'        End With
'    Next sheet
'
'    ' 処理したシートを選択し、PDFとして保存
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' シートの元の設定に戻す
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' 元の回転角度に戻す
'        If ws.Shapes.Count > 0 Then
'            Set shapeGroup = ws.Shapes(ws.Shapes.Count)
'            shapeGroup.Rotation = originalRotations(i)
'        End If
'
'        ' 元のページ設定に戻す
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'
'        i = i + 1
'    Next sheet
'
'    MsgBox "指定されたシートが回転され、PDFとして保存されました。その後、元の状態に復元されました。"
'End Sub

'Sub RotateAndExportSheetsToPDF(sheetNames As Collection, outputFileName As String)
'    Dim ws As Worksheet
'    Dim shapeArray() As Variant
'    Dim shp As shape
'    Dim shapeGroup As shape
'    Dim i As Integer
'    Dim selectedSheets() As String
'    Dim pdfPath As String
'    Dim sheet As Variant
'    Dim originalRotations As Collection
'    Dim originalOrientations As Collection
'    Dim originalZoom As Collection
'
'    ' 各シートの元の設定を保存するためのコレクションを初期化
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF保存先のパスを指定
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' 引数として受け取ったシート名のコレクションを配列に変換
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' 各シートを処理
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' シート上のすべてのシェイプをグループ化して90度回転
'        If ws.Shapes.Count > 0 Then
'            ReDim shapeArray(1 To ws.Shapes.Count)
'            For i = 1 To ws.Shapes.Count
'                shapeArray(i) = ws.Shapes(i).Name
'            Next i
'
'            Set shapeGroup = ws.Shapes.Range(shapeArray).Group
'
'            ' 元の回転角度を保存
'            originalRotations.Add shapeGroup.Rotation
'
'            ' 90度回転
'            shapeGroup.Rotation = shapeGroup.Rotation + 90
'
'            ' グループがシートの範囲を超えないように再配置
'            If shapeGroup.left < 0 Then shapeGroup.left = 0
'            If shapeGroup.top < 0 Then shapeGroup.top = 0
'        End If
'
'        ' 元のページ設定を保存
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' ページ設定で縮尺を無効にして設定
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' または xlLandscape、必要に応じて設定
'        End With
'    Next sheet
'
'    ' 処理したシートを選択し、PDFとして保存
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' シートの元の設定に戻す
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' グループを解除して個別のシェイプに回転を適用
'        If ws.Shapes.Count > 0 Then
'            ' グループを解除
'            ws.Shapes(ws.Shapes.Count).Ungroup
'
'            ' 個別のシェイプに元の回転角度を設定
'            For Each shp In ws.Shapes
'                shp.Rotation = originalRotations(i)
'            Next shp
'        End If
'
'        ' 元のページ設定に戻す
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'
'        i = i + 1
'    Next sheet
'
'    MsgBox "指定されたシートが回転され、PDFとして保存されました。その後、元の状態に復元されました。"
'End Sub

'Sub RotateAndExportSheetsToPDF(sheetNames As Collection, outputFileName As String)
'    Dim ws As Worksheet
'    Dim shapeArray() As Variant
'    Dim shp As shape
'    Dim shapeGroup As shape
'    Dim i As Integer
'    Dim selectedSheets() As String
'    Dim pdfPath As String
'    Dim sheet As Variant
'    Dim originalRotations As Collection
'    Dim originalOrientations As Collection
'    Dim originalZoom As Collection
'
'    ' 各シートの元の設定を保存するためのコレクションを初期化
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF保存先のパスを指定
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' 引数として受け取ったシート名のコレクションを配列に変換
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' 各シートを処理
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' シート上のすべてのシェイプをグループ化して90度回転
'        If ws.Shapes.Count > 0 Then
'            ReDim shapeArray(1 To ws.Shapes.Count)
'            For i = 1 To ws.Shapes.Count
'                shapeArray(i) = ws.Shapes(i).Name
'            Next i
'
'            Set shapeGroup = ws.Shapes.Range(shapeArray).Group
'
'            ' 元の回転角度を保存
'            originalRotations.Add shapeGroup.Rotation
'
'            ' 90度回転
'            shapeGroup.Rotation = shapeGroup.Rotation + 90
'
'            ' グループがシートの範囲を超えないように再配置
'            If shapeGroup.left < 0 Then shapeGroup.left = 0
'            If shapeGroup.top < 0 Then shapeGroup.top = 0
'        End If
'
'        ' 元のページ設定を保存
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' ページ設定で縮尺を無効にして設定
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' または xlLandscape、必要に応じて設定
'        End With
'    Next sheet
'
'    ' 処理したシートを選択し、PDFとして保存
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' シートの元の設定に戻す
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' グループを解除して個別のシェイプに回転を適用
'        If ws.Shapes.Count > 0 Then
'            ' グループを解除
'            shapeGroup.Ungroup
'
'            ' 個別のシェイプに元の回転角度を設定
'            For Each shp In ws.Shapes
'                shp.Rotation = originalRotations(i)
'            Next shp
'        End If
'
'        ' 元のページ設定に戻す
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'
'        i = i + 1
'    Next sheet
'
'    MsgBox "指定されたシートが回転され、PDFとして保存されました。その後、元の状態に復元されました。"
'End Sub
'Sub RotateAndExportSheetsToPDF(sheetNames As Collection, outputFileName As String)
'    Dim ws As Worksheet
'    Dim shp As shape
'    Dim shapeArray() As Variant
'    Dim i As Integer
'    Dim selectedSheets() As String
'    Dim pdfPath As String
'    Dim sheet As Variant
'    Dim originalRotations As Collection
'    Dim originalOrientations As Collection
'    Dim originalZoom As Collection
'
'    ' 各シートの元の設定を保存するためのコレクションを初期化
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF保存先のパスを指定
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' 引数として受け取ったシート名のコレクションを配列に変換
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' 各シートを処理
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' 各シート上のシェイプの回転角度を保存し、90度回転させる
'        If ws.Shapes.Count > 0 Then
'            For Each shp In ws.Shapes
'                originalRotations.Add shp.Rotation
'                shp.Rotation = shp.Rotation + 90
'            Next shp
'        End If
'
'        ' 元のページ設定を保存
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' ページ設定で縮尺を無効にして設定
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' または xlLandscape、必要に応じて設定
'        End With
'    Next sheet
'
'    ' 処理したシートを選択し、PDFとして保存
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' シートの元の設定に戻す
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' 個別のシェイプに元の回転角度を設定
'        If ws.Shapes.Count > 0 Then
'            For Each shp In ws.Shapes
'                shp.Rotation = originalRotations(i)
'                i = i + 1
'            Next shp
'        End If
'
'        ' 元のページ設定に戻す
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'    Next sheet
'
'    MsgBox "指定されたシートが回転され、PDFとして保存されました。その後、元の状態に復元されました。"
'End Sub

Sub ExportSheetsToPDFWithoutRotation(sheetNames As Collection, outputFileName As String)
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim i As Integer
    Dim selectedSheets() As String
    Dim sheet As Variant
    
    ' PDF保存先のパスを指定
    pdfPath = ThisWorkbook.Path & "\"
    
    ' 引数として受け取ったシート名のコレクションを配列に変換
    ReDim selectedSheets(1 To sheetNames.Count)
    i = 1
    For Each sheet In sheetNames
        selectedSheets(i) = sheet
        i = i + 1
    Next sheet
    
    ' 各シートのページ設定を調整
    For Each sheet In sheetNames
        Set ws = ThisWorkbook.Sheets(sheet)
        
        With ws.PageSetup
            .Zoom = False
            .FitToPagesWide = 1 ' シートを1ページに収める
            .FitToPagesTall = 1 ' シートを1ページに収める
            '.Orientation = xlLandscape ' 必要に応じて縦向きや横向きを設定
            .Orientation = xlPortrait
        End With
    Next sheet
    
    ' 選択したシートをPDFとして保存
    ThisWorkbook.Sheets(selectedSheets).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfPath & outputFileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    MsgBox "指定されたシートがPDFとして " & pdfPath & outputFileName & " に保存されました。"
End Sub

' 使用例
Sub ExampleUsageRotateAndExport()
    Dim sheetNames As New Collection
    sheetNames.Add "classB_1F"
    sheetNames.Add "classA_2F"
    sheetNames.Add "classA_1F"
    
    ExportSheetsToPDFWithoutRotation sheetNames, "RotatedSheets.pdf"
End Sub


