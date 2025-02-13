Attribute VB_Name = "模块1"
Function CallTongyiAPI(api_key As String, inputText As String) As String
    Dim API As String
    Dim SendTxt As String
    Dim Http As Object
    Dim status_code As Integer
    Dim response As String
 '该请求为http请求，应该将此网址替换成对应的URL
    API = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    SendTxt = "{""model"": ""qwen-max-2025-01-25"", ""messages"": [{""role"":""system"", ""content"":""You are a Word assistant""}, {""role"":""user"", ""content"":""" & inputText & """}], ""stream"": false}"
 'qwen-max-2025-01-25 为模型别名，接入api时需要修改
 'stream建议设置成false，之前试过true，生成内容无法返回
    Set Http = CreateObject("MSXML2.XMLHTTP")
    With Http
      .Open "POST", API, False
      .setRequestHeader "Content-Type", "application/json"
      .setRequestHeader "Authorization", "Bearer " & api_key
      .send SendTxt
        status_code = .Status
        response = .responseText
    End With

    If status_code = 200 Then
        CallTongyiAPI = response
    Else
        CallTongyiAPI = "Error: " & status_code & " - " & response
    End If

    Set Http = Nothing
End Function

Sub TongyiV3()
    Dim api_key As String
    Dim inputText As String
    Dim response As String
    Dim regex As Object
    Dim matches As Object
    Dim originalSelection As Object

    api_key = "" ' 替换为你的api key
    If api_key = "" Then
        MsgBox "Please enter the API key."
        Exit Sub
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select text."
        Exit Sub
    End If

    ' 保存原始选中的文本
    Set originalSelection = Selection.Range.Duplicate

    inputText = Replace(Replace(Replace(Replace(Replace(Selection.Text, "\", "\\"), vbCrLf, ""), vbCr, ""), vbLf, ""), Chr(34), "\""")
    response = CallTongyiAPI(api_key, inputText)

    If Left(response, 5) <> "Error" Then
        Set regex = CreateObject("VBScript.RegExp")
        With regex
          .Global = True
          .MultiLine = True
          .IgnoreCase = False
          .Pattern = """content"":""(.*?)"""
        End With
        Set matches = regex.Execute(response)
        If matches.Count > 0 Then
            response = matches(0).SubMatches(0)
            response = Replace(Replace(response, """", Chr(34)), """", Chr(34))
            ' 将 \n 替换为 Word 中的段落标记 vbCrLf
            response = Replace(response, "\n", vbCrLf)

            ' 取消选中原始文本
            Selection.Collapse Direction:=wdCollapseEnd

            ' 将内容插入到选中文字的下一行
            Selection.TypeParagraph ' 插入新行
            Selection.TypeText Text:=response

            ' 将光标移回原来选中文本的末尾
            originalSelection.Select
        Else
            MsgBox "Failed to parse API response.", vbExclamation
        End If
    Else
        MsgBox response, vbCritical
    End If
End Sub
