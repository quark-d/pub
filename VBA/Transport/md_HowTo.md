Sub WriteLargeDataUsingTranspose()
    Const MAX_CHUNK_SIZE As Long = 65536
    Const START_ROW As Long = 6
    Const TOTAL_ROWS As Long = 1048576 ' Excel最大行

    Dim dataArr() As Variant
    Dim chunkArr() As Variant
    Dim i As Long, r As Long
    Dim chunkStart As Long, chunkEnd As Long
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")

    ' データ作成
    ReDim dataArr(START_ROW To TOTAL_ROWS)
    For i = START_ROW To TOTAL_ROWS
        dataArr(i) = "データ" & i
    Next i

    ' チャンクごとに書き込み
    r = 0
    Do While START_ROW + r <= TOTAL_ROWS
        chunkStart = START_ROW + r
        chunkEnd = Application.Min(chunkStart + MAX_CHUNK_SIZE - 1, TOTAL_ROWS)
        
        ReDim chunkArr(chunkStart To chunkEnd)
        For i = chunkStart To chunkEnd
            chunkArr(i) = dataArr(i)
        Next i

        ws.Range("A" & chunkStart).Resize(chunkEnd - chunkStart + 1, 1).Value = Application.Transpose(chunkArr)
        r = r + MAX_CHUNK_SIZE
    Loop
End Sub