
## url
```vb
Function GetCache_Products() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT link_name, link_url FROM dbo_tbl_Products", dbOpenSnapshot)
    
    Do Until rs.EOF
        dict(rs!link_name) = Nz(rs!link_url, "")
        rs.MoveNext
    Loop
    
    rs.Close
    Set GetCache_Products = dict
End Function
```

## comment
```vb
Function GetCache_ProductsB() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT link_name, link_comment FROM dbo_tbl_ProductsB", dbOpenSnapshot)
    
    Do Until rs.EOF
        dict(rs!link_name) = Nz(rs!link_comment, "")
        rs.MoveNext
    Loop
    
    rs.Close
    Set GetCache_ProductsB = dict
End Function
```

## type
```vb
Function GetCache_ProductsC() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT link_name, link_type FROM dbo_tbl_ProductsC", dbOpenSnapshot)
    
    Do Until rs.EOF
        dict(rs!link_name) = Nz(rs!link_type, "")
        rs.MoveNext
    Loop
    
    rs.Close
    Set GetCache_ProductsC = dict
End Function
```

## 差分を検出して更新・追加
```vba
Sub UpdateBulk_All(linkData As Collection)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim item As Variant
    Dim linkName As String, linkUrl As String, linkComment As String, linkType As String
    
    ' 一括キャッシュ取得（SQL Serverへの接続は各1回）
    Dim cacheUrl As Object: Set cacheUrl = GetCache_Products()
    Dim cacheComment As Object: Set cacheComment = GetCache_ProductsB()
    Dim cacheType As Object: Set cacheType = GetCache_ProductsC()
    
    For Each item In linkData
        linkName = item(0)
        linkUrl = item(1)
        linkComment = item(2)
        linkType = item(3)

        ' -------- URL（dbo_tbl_Products）
        If Not cacheUrl.exists(linkName) Then
            Set rs = db.OpenRecordset("dbo_tbl_Products", dbOpenDynaset)
            rs.AddNew
            rs!link_name = linkName
            rs!link_url = linkUrl
            rs.Update
            rs.Close
        ElseIf cacheUrl(linkName) <> linkUrl Then
            Set rs = db.OpenRecordset("SELECT * FROM dbo_tbl_Products WHERE link_name = '" & Replace(linkName, "'", "''") & "'", dbOpenDynaset)
            If Not rs.EOF Then
                rs.Edit
                rs!link_url = linkUrl
                rs.Update
            End If
            rs.Close
        End If

        ' -------- コメント（dbo_tbl_ProductsB）
        If Not cacheComment.exists(linkName) Then
            Set rs = db.OpenRecordset("dbo_tbl_ProductsB", dbOpenDynaset)
            rs.AddNew
            rs!link_name = linkName
            rs!link_comment = linkComment
            rs.Update
            rs.Close
        ElseIf cacheComment(linkName) <> linkComment Then
            Set rs = db.OpenRecordset("SELECT * FROM dbo_tbl_ProductsB WHERE link_name = '" & Replace(linkName, "'", "''") & "'", dbOpenDynaset)
            If Not rs.EOF Then
                rs.Edit
                rs!link_comment = linkComment
                rs.Update
            End If
            rs.Close
        End If

        ' -------- タイプ（dbo_tbl_ProductsC）
        If Not cacheType.exists(linkName) Then
            Set rs = db.OpenRecordset("dbo_tbl_ProductsC", dbOpenDynaset)
            rs.AddNew
            rs!link_name = linkName
            rs!link_type = linkType
            rs.Update
            rs.Close
        ElseIf cacheType(linkName) <> linkType Then
            Set rs = db.OpenRecordset("SELECT * FROM dbo_tbl_ProductsC WHERE link_name = '" & Replace(linkName, "'", "''") & "'", dbOpenDynaset)
            If Not rs.EOF Then
                rs.Edit
                rs!link_type = linkType
                rs.Update
            End If
            rs.Close
        End If
    Next item
End Sub
```

## 呼び出し例
```vb
Sub ExampleCall_All()
    Dim data As New Collection
    data.Add Array("link1", "https://url1.com", "コメント1", "タイプA")
    data.Add Array("link2", "https://url2.com", "コメント2", "タイプB")
    data.Add Array("link3", "https://url3.com", "コメント3", "タイプC")
    
    Call UpdateBulk_All(data)
End Sub
```


