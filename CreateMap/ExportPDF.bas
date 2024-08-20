Attribute VB_Name = "ExportPDF"
Option Explicit

Public Sub ExportSheetsWithColorTabsToPDF(Optional colorIndex As Long = 3)
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Set sheetNames = New Collection
    Dim pdfPath As String
    
    ' �w�肳�ꂽ�F�̃^�u�����V�[�g�����X�g�ɒǉ�
    For Each ws In ThisWorkbook.Sheets
        If ws.Tab.colorIndex = colorIndex Then
            sheetNames.Add ws.Name
        End If
    Next ws
    
    ' ����Ώۂ̃V�[�g�����邩�ǂ������m�F
    If sheetNames.Count > 0 Then
        ' ���X�g�������Ƃ��ēn����PDF�ɃG�N�X�|�[�g
        pdfPath = ExportSheetsToPDFWithoutRotation(sheetNames, "output.pdf")
        
        ' PDF�t�@�C�����J��
        ThisWorkbook.FollowHyperlink pdfPath
    Else
        MsgBox "����Ώۂ̃V�[�g�͂���܂���B", vbInformation
    End If
End Sub



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
'    ' �e�V�[�g�̌��̐ݒ��ۑ����邽�߂̃R���N�V������������
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF�ۑ���̃p�X���w��
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' �����Ƃ��Ď󂯎�����V�[�g���̃R���N�V������z��ɕϊ�
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' �e�V�[�g������
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �V�[�g��̂��ׂẴV�F�C�v���O���[�v������90�x��]
'        If ws.Shapes.Count > 0 Then
'            ReDim shapeArray(1 To ws.Shapes.Count)
'            For i = 1 To ws.Shapes.Count
'                shapeArray(i) = ws.Shapes(i).Name
'            Next i
'
'            Set shapeGroup = ws.Shapes.Range(shapeArray).Group
'
'            ' ���̉�]�p�x��ۑ�
'            originalRotations.Add shapeGroup.Rotation
'
'            ' 90�x��]
'            shapeGroup.Rotation = shapeGroup.Rotation + 90
'
'            ' �O���[�v���V�[�g�͈̔͂𒴂��Ȃ��悤�ɍĔz�u
'            If shapeGroup.left < 0 Then shapeGroup.left = 0
'            If shapeGroup.top < 0 Then shapeGroup.top = 0
'        End If
'
'        ' ���̃y�[�W�ݒ��ۑ�
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' �y�[�W�ݒ�ŏk�ڂ𖳌��ɂ��Đݒ�
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' �܂��� xlLandscape�A�K�v�ɉ����Đݒ�
'        End With
'    Next sheet
'
'    ' ���������V�[�g��I�����APDF�Ƃ��ĕۑ�
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' �V�[�g�̌��̐ݒ�ɖ߂�
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' ���̉�]�p�x�ɖ߂�
'        If ws.Shapes.Count > 0 Then
'            Set shapeGroup = ws.Shapes(ws.Shapes.Count)
'            shapeGroup.Rotation = originalRotations(i)
'        End If
'
'        ' ���̃y�[�W�ݒ�ɖ߂�
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'
'        i = i + 1
'    Next sheet
'
'    MsgBox "�w�肳�ꂽ�V�[�g����]����APDF�Ƃ��ĕۑ�����܂����B���̌�A���̏�Ԃɕ�������܂����B"
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
'    ' �e�V�[�g�̌��̐ݒ��ۑ����邽�߂̃R���N�V������������
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF�ۑ���̃p�X���w��
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' �����Ƃ��Ď󂯎�����V�[�g���̃R���N�V������z��ɕϊ�
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' �e�V�[�g������
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �V�[�g��̂��ׂẴV�F�C�v���O���[�v������90�x��]
'        If ws.Shapes.Count > 0 Then
'            ReDim shapeArray(1 To ws.Shapes.Count)
'            For i = 1 To ws.Shapes.Count
'                shapeArray(i) = ws.Shapes(i).Name
'            Next i
'
'            Set shapeGroup = ws.Shapes.Range(shapeArray).Group
'
'            ' ���̉�]�p�x��ۑ�
'            originalRotations.Add shapeGroup.Rotation
'
'            ' 90�x��]
'            shapeGroup.Rotation = shapeGroup.Rotation + 90
'
'            ' �O���[�v���V�[�g�͈̔͂𒴂��Ȃ��悤�ɍĔz�u
'            If shapeGroup.left < 0 Then shapeGroup.left = 0
'            If shapeGroup.top < 0 Then shapeGroup.top = 0
'        End If
'
'        ' ���̃y�[�W�ݒ��ۑ�
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' �y�[�W�ݒ�ŏk�ڂ𖳌��ɂ��Đݒ�
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' �܂��� xlLandscape�A�K�v�ɉ����Đݒ�
'        End With
'    Next sheet
'
'    ' ���������V�[�g��I�����APDF�Ƃ��ĕۑ�
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' �V�[�g�̌��̐ݒ�ɖ߂�
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �O���[�v���������Čʂ̃V�F�C�v�ɉ�]��K�p
'        If ws.Shapes.Count > 0 Then
'            ' �O���[�v������
'            ws.Shapes(ws.Shapes.Count).Ungroup
'
'            ' �ʂ̃V�F�C�v�Ɍ��̉�]�p�x��ݒ�
'            For Each shp In ws.Shapes
'                shp.Rotation = originalRotations(i)
'            Next shp
'        End If
'
'        ' ���̃y�[�W�ݒ�ɖ߂�
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'
'        i = i + 1
'    Next sheet
'
'    MsgBox "�w�肳�ꂽ�V�[�g����]����APDF�Ƃ��ĕۑ�����܂����B���̌�A���̏�Ԃɕ�������܂����B"
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
'    ' �e�V�[�g�̌��̐ݒ��ۑ����邽�߂̃R���N�V������������
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF�ۑ���̃p�X���w��
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' �����Ƃ��Ď󂯎�����V�[�g���̃R���N�V������z��ɕϊ�
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' �e�V�[�g������
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �V�[�g��̂��ׂẴV�F�C�v���O���[�v������90�x��]
'        If ws.Shapes.Count > 0 Then
'            ReDim shapeArray(1 To ws.Shapes.Count)
'            For i = 1 To ws.Shapes.Count
'                shapeArray(i) = ws.Shapes(i).Name
'            Next i
'
'            Set shapeGroup = ws.Shapes.Range(shapeArray).Group
'
'            ' ���̉�]�p�x��ۑ�
'            originalRotations.Add shapeGroup.Rotation
'
'            ' 90�x��]
'            shapeGroup.Rotation = shapeGroup.Rotation + 90
'
'            ' �O���[�v���V�[�g�͈̔͂𒴂��Ȃ��悤�ɍĔz�u
'            If shapeGroup.left < 0 Then shapeGroup.left = 0
'            If shapeGroup.top < 0 Then shapeGroup.top = 0
'        End If
'
'        ' ���̃y�[�W�ݒ��ۑ�
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' �y�[�W�ݒ�ŏk�ڂ𖳌��ɂ��Đݒ�
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' �܂��� xlLandscape�A�K�v�ɉ����Đݒ�
'        End With
'    Next sheet
'
'    ' ���������V�[�g��I�����APDF�Ƃ��ĕۑ�
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' �V�[�g�̌��̐ݒ�ɖ߂�
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �O���[�v���������Čʂ̃V�F�C�v�ɉ�]��K�p
'        If ws.Shapes.Count > 0 Then
'            ' �O���[�v������
'            shapeGroup.Ungroup
'
'            ' �ʂ̃V�F�C�v�Ɍ��̉�]�p�x��ݒ�
'            For Each shp In ws.Shapes
'                shp.Rotation = originalRotations(i)
'            Next shp
'        End If
'
'        ' ���̃y�[�W�ݒ�ɖ߂�
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'
'        i = i + 1
'    Next sheet
'
'    MsgBox "�w�肳�ꂽ�V�[�g����]����APDF�Ƃ��ĕۑ�����܂����B���̌�A���̏�Ԃɕ�������܂����B"
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
'    ' �e�V�[�g�̌��̐ݒ��ۑ����邽�߂̃R���N�V������������
'    Set originalRotations = New Collection
'    Set originalOrientations = New Collection
'    Set originalZoom = New Collection
'
'    ' PDF�ۑ���̃p�X���w��
'    pdfPath = ThisWorkbook.Path & "\"
'
'    ' �����Ƃ��Ď󂯎�����V�[�g���̃R���N�V������z��ɕϊ�
'    ReDim selectedSheets(1 To sheetNames.Count)
'    i = 1
'    For Each sheet In sheetNames
'        selectedSheets(i) = sheet
'        i = i + 1
'    Next sheet
'
'    ' �e�V�[�g������
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �e�V�[�g��̃V�F�C�v�̉�]�p�x��ۑ����A90�x��]������
'        If ws.Shapes.Count > 0 Then
'            For Each shp In ws.Shapes
'                originalRotations.Add shp.Rotation
'                shp.Rotation = shp.Rotation + 90
'            Next shp
'        End If
'
'        ' ���̃y�[�W�ݒ��ۑ�
'        originalOrientations.Add ws.PageSetup.Orientation
'        originalZoom.Add ws.PageSetup.Zoom
'
'        ' �y�[�W�ݒ�ŏk�ڂ𖳌��ɂ��Đݒ�
'        With ws.PageSetup
'            .Zoom = False
'            .FitToPagesWide = 1
'            .FitToPagesTall = 1
'            .Orientation = xlPortrait ' �܂��� xlLandscape�A�K�v�ɉ����Đݒ�
'        End With
'    Next sheet
'
'    ' ���������V�[�g��I�����APDF�Ƃ��ĕۑ�
'    ThisWorkbook.Sheets(selectedSheets).Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
'        Filename:=pdfPath & outputFileName, _
'        Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, _
'        IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False
'
'    ' �V�[�g�̌��̐ݒ�ɖ߂�
'    i = 1
'    For Each sheet In sheetNames
'        Set ws = ThisWorkbook.Sheets(sheet)
'
'        ' �ʂ̃V�F�C�v�Ɍ��̉�]�p�x��ݒ�
'        If ws.Shapes.Count > 0 Then
'            For Each shp In ws.Shapes
'                shp.Rotation = originalRotations(i)
'                i = i + 1
'            Next shp
'        End If
'
'        ' ���̃y�[�W�ݒ�ɖ߂�
'        With ws.PageSetup
'            .Orientation = originalOrientations(i)
'            .Zoom = originalZoom(i)
'        End With
'    Next sheet
'
'    MsgBox "�w�肳�ꂽ�V�[�g����]����APDF�Ƃ��ĕۑ�����܂����B���̌�A���̏�Ԃɕ�������܂����B"
'End Sub
Function ExportSheetsToPDFWithoutRotation(sheetNames As Collection, outputFileName As String) As String
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim i As Integer
    Dim selectedSheets() As String
    Dim sheet As Variant
    
    ' PDF�ۑ���̃p�X���w��
    pdfPath = ThisWorkbook.Path & "\" & outputFileName
    
    ' �����Ƃ��Ď󂯎�����V�[�g���̃R���N�V������z��ɕϊ�
    ReDim selectedSheets(1 To sheetNames.Count)
    i = 1
    For Each sheet In sheetNames
        selectedSheets(i) = sheet
        i = i + 1
    Next sheet
    
    ' �e�V�[�g�̃y�[�W�ݒ�𒲐�
    For Each sheet In sheetNames
        Set ws = ThisWorkbook.Sheets(sheet)
        
        With ws.PageSetup
            .Zoom = False
            .FitToPagesWide = 1 ' �V�[�g��1�y�[�W�Ɏ��߂�
            .FitToPagesTall = 1 ' �V�[�g��1�y�[�W�Ɏ��߂�
            .Orientation = xlLandscape ' �K�v�ɉ����ďc�����≡������ݒ�
            ' .Orientation = xlPortrait
        End With
    Next sheet
    
    ' �e�V�[�g�̃w�b�_�[�ɃV�[�g����ݒ�i�����ŕ\���j
    For Each sheet In sheetNames
        Set ws = ThisWorkbook.Sheets(sheet)
        ws.PageSetup.CenterHeader = "&B" & ws.Name ' ����(&B)�Ńw�b�_�[�ɃV�[�g����ݒ�
    Next sheet
    
    ' �I�������V�[�g��PDF�Ƃ��ĕۑ�
    ThisWorkbook.Sheets(selectedSheets).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    MsgBox "�w�肳�ꂽ�V�[�g��PDF�Ƃ��� " & pdfPath & " �ɕۑ�����܂����B"
    
    ' �����Ƀw�b�_�[���N���A
    For Each sheet In sheetNames
        Set ws = ThisWorkbook.Sheets(sheet)
        ws.PageSetup.CenterHeader = "" ' �w�b�_�[���N���A
    Next sheet
    
    ExportSheetsToPDFWithoutRotation = pdfPath
End Function

' �g�p��
Sub ExampleUsageRotateAndExport()
    Dim sheetNames As New Collection
    sheetNames.Add "classB_1F"
    sheetNames.Add "classA_2F"
    sheetNames.Add "classA_1F"
    
    ExportSheetsToPDFWithoutRotation sheetNames, "RotatedSheets.pdf"
End Sub


