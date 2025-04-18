# 制限

https://newtonexcelbach.com/2016/01/01/worksheetfunction-transpose-changed-behaviour-in-excel-2013-and-2016/


## チャンクして対応
```vb
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

## 転置を自前で用意する
```vb
Sub TransposeArray()
    Dim sourceArray() As Variant
    Dim transposedArray() As Variant
    Dim i As Long
    Dim lastRow As Long

    ' データの作成（例：6行目から最終行まで）
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then lastRow = 6
    ReDim sourceArray(6 To lastRow)
    For i = 6 To lastRow
        sourceArray(i) = "データ" & i
    Next i

    ' 転置処理
    ReDim transposedArray(1 To 1, 1 To lastRow - 5)
    For i = 6 To lastRow
        transposedArray(1, i - 5) = sourceArray(i)
    Next i

    ' シートへの書き込み（例：A6セルから）
    Range("A6").Resize(1, lastRow - 5).Value = transposedArray
End Sub
```