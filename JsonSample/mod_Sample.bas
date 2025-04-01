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

    ' JSON�f�[�^�i���ڃR�[�h�ɖ��ߍ��݁j
    jsonText = "{" & _
        """team"": {" & _
            """name"": ""teamA""," & _
            """leader"": {""name"": ""Manager"", ""age"": 25, ""note"": ""�J���`�[���̃��[�_�[""}," & _
            """member"": [" & _
                "{""name"": ""Alice"", ""age"": 25, ""note"": ""�J���`�[����\n�T�u���[�_�[""}," & _
                "{""name"": ""Bob"", ""age"": 25, ""note"": ""�J���`�[���̃��[�_�[""}" & _
            "]" & _
        "}" & _
    "}"

    ' JSON���
    Set parsed = JsonConverter.ParseJson(jsonText)
    Set team = parsed("team")
    Set leader = team("leader")
    Set members = team("member")

    ' �`�[����
    Debug.Print "���`�[����: " & team("name")

    ' ���[�_�[���
    Debug.Print "�����[�_�[: " & leader("name") & "�i�N��: " & leader("age") & "�j"
    Debug.Print "���l: " & leader("note")

    ' �����o�[���
    Debug.Print "�������o�[�ꗗ:"
    For Each member In members
        Debug.Print "���O: " & member("name") & ", �N��: " & member("age") & ", ���l: " & member("note")
    Next member
End Sub

Sub FetchJsonFromWeb()
    Dim http As Object
    Dim jsonText As String
    Dim result As Collection
    Dim user As Variant

    ' HTTP���N�G�X�g�p�I�u�W�F�N�g
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' JSONPlaceholder����f�[�^�擾
    http.Open "GET", "https://faopfjiaopweijf.free.beeceptor.com/test02", False
    http.Send
    
    ' ���X�|���X��JSON������
    jsonText = http.responseText
    Debug.Print jsonText
    
    ' JSON��́i�z��ɂȂ�j
    Set result = JsonConverter.ParseJson(jsonText)
    
    ' ���ʂ�\��
    For Each user In result
        Debug.Print "ID: " & user("id") & _
                    ", ���O: " & user("name") & _
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

    ' HTTP�ʐM
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://faopfjiaopweijf.free.beeceptor.com/test03"
    http.Send
    jsonText = http.responseText
    Debug.Print jsonText

    ' �S�̂�JSON���p�[�X
    Set result = JsonConverter.ParseJson(jsonText)
    Set team = result("team")

    ' ���[�_�[���
    Set leader = team("leader")
    Debug.Print "�`�[��: " & team("name")
    Debug.Print "���[�_�[: " & leader("name") & "�i" & leader("age") & "�΁j"

    ' �����o�[��JSON����������o���āA�ăp�[�X
    memberJson = team("member")
    Debug.Print memberJson
    Set members = JsonConverter.ParseJson(memberJson)

    ' �����o�[�ꗗ
    For Each m In members
        Debug.Print "���O: " & m("name") & ", �N��: " & m("age") & ", ���l: " & m("note")
    Next m
End Sub
