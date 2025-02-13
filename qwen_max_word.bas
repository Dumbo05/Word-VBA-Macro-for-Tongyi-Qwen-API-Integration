Attribute VB_Name = "ģ��1"
Function CallTongyiAPI(api_key As String, inputText As String) As String
    Dim API As String
    Dim SendTxt As String
    Dim Http As Object
    Dim status_code As Integer
    Dim response As String
 '������Ϊhttp����Ӧ�ý�����ַ�滻�ɶ�Ӧ��URL
    API = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    SendTxt = "{""model"": ""qwen-max-2025-01-25"", ""messages"": [{""role"":""system"", ""content"":""You are a Word assistant""}, {""role"":""user"", ""content"":""" & inputText & """}], ""stream"": false}"
 'qwen-max-2025-01-25 Ϊģ�ͱ���������apiʱ��Ҫ�޸�
 'stream�������ó�false��֮ǰ�Թ�true�����������޷�����
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

    api_key = "sk-df73901a945f4247b064a7d078f5edfe" ' �滻Ϊ���api key
    If api_key = "" Then
        MsgBox "Please enter the API key."
        Exit Sub
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select text."
        Exit Sub
    End If

    ' ����ԭʼѡ�е��ı�
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
            ' �� \n �滻Ϊ Word �еĶ����� vbCrLf
            response = Replace(response, "\n", vbCrLf)

            ' ȡ��ѡ��ԭʼ�ı�
            Selection.Collapse Direction:=wdCollapseEnd

            ' �����ݲ��뵽ѡ�����ֵ���һ��
            Selection.TypeParagraph ' ��������
            Selection.TypeText Text:=response

            ' ������ƻ�ԭ��ѡ���ı���ĩβ
            originalSelection.Select
        Else
            MsgBox "Failed to parse API response.", vbExclamation
        End If
    Else
        MsgBox response, vbCritical
    End If
End Sub
