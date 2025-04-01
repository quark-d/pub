Attribute VB_Name = "mod_Sample"
Option Explicit

Sub ParseTeamJson()
    Dim jsonText As String
    ' Microsoft Scripting Runtime
    Dim parsed As Dictionary
    Dim team As Dictionary
    Dim leader As Dictionary
    Dim members As Collection
    Dim member As Variant

    ' JSONデータ（直接コードに埋め込み）
    jsonText = "{" & _
        """team"": {" & _
            """name"": ""teamA""," & _
            """leader"": {""name"": ""Manager"", ""age"": 25, ""note"": ""開発チームのリーダー""}," & _
            """member"": [" & _
                "{""name"": ""Alice"", ""age"": 25, ""note"": ""開発チームの\nサブリーダー""}," & _
                "{""name"": ""Bob"", ""age"": 25, ""note"": ""開発チームのリーダー""}" & _
            "]" & _
        "}" & _
    "}"

    ' JSON解析
    Set parsed = JsonConverter.ParseJson(jsonText)
    Set team = parsed("team")
    Set leader = team("leader")
    Set members = team("member")

    ' チーム名
    Debug.Print "■チーム名: " & team("name")

    ' リーダー情報
    Debug.Print "●リーダー: " & leader("name") & "（年齢: " & leader("age") & "）"
    Debug.Print "備考: " & leader("note")

    ' メンバー情報
    Debug.Print "●メンバー一覧:"
    For Each member In members
        Debug.Print "名前: " & member("name") & ", 年齢: " & member("age") & ", 備考: " & member("note")
    Next member
End Sub

Sub FetchJsonFromWeb()
    Dim http As Object
    Dim jsonText As String
    Dim result As Collection
    Dim user As Variant

    ' HTTPリクエスト用オブジェクト
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' JSONPlaceholderからデータ取得
    http.Open "GET", "https://faopfjiaopweijf.free.beeceptor.com/test02", False
    http.Send
    
    ' レスポンスのJSON文字列
    jsonText = http.responseText
    Debug.Print jsonText
    
    ' JSON解析（配列になる）
    Set result = JsonConverter.ParseJson(jsonText)
    
    ' 結果を表示
    For Each user In result
        Debug.Print "ID: " & user("id") & _
                    ", 名前: " & user("name") & _
                    ", Email: " & user("email")
    Next user
End Sub

Sub FetchAndParseNestedJsonString()
    Dim http As Object
    Dim jsonText As String
    Dim result As Dictionary
    Dim team As Dictionary
    Dim leader As Dictionary
    Dim memberJson As String
    Dim members As Collection
    Dim m As Variant

    ' HTTP通信
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://faopfjiaopweijf.free.beeceptor.com/test03"
    http.Send
    jsonText = http.responseText
    Debug.Print jsonText

    ' 全体のJSONをパース
    Set result = JsonConverter.ParseJson(jsonText)
    Set team = result("team")

    ' リーダー情報
    Set leader = team("leader")
    Debug.Print "チーム: " & team("name")
    Debug.Print "リーダー: " & leader("name") & "（" & leader("age") & "歳）"

    ' メンバーのJSON文字列を取り出して、再パース
    memberJson = team("member")
    Debug.Print memberJson
    Set members = JsonConverter.ParseJson(memberJson)

    ' メンバー一覧
    For Each m In members
        Debug.Print "名前: " & m("name") & ", 年齢: " & m("age") & ", 備考: " & m("note")
    Next m
End Sub
