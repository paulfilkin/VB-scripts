Attribute VB_Name = "CallOpenAI"
Function CallOpenAI_prompt(userMessage As String, template As String) As String
    On Error GoTo ErrorHandler

    Dim xmlHttp As Object
    Dim strUrl As String
    Dim strPayload As String
    Dim strResponse As String
    Dim jsonResponse As Object
    Dim API_KEY As String
    Dim model As String
    Dim ws As Worksheet
    
    ' Attempt to set the worksheet and retrieve the API key and model
    On Error GoTo SheetError
    Set ws = ThisWorkbook.Sheets("Key_models")
    API_KEY = ws.Range("E3").Value
    model = ws.Range("E5").Value
    On Error GoTo ErrorHandler

    ' Escape quotes in the userMessage and template
    userMessage = Replace(userMessage, """", "\""")
    template = Replace(template, """", "\""")
    
    ' Define the endpoint and the payload
    strUrl = "https://api.openai.com/v1/chat/completions"
    strPayload = "{""model"": """ & model & """, ""messages"": [{""role"": ""system"", ""content"": """ & template & """}, {""role"": ""user"", ""content"": """ & userMessage & """}]}"
    
    ' Create an XML HTTP request
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Open the connection and set headers
    xmlHttp.Open "POST", strUrl, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Authorization", "Bearer " & API_KEY
    
    ' Send the request
    xmlHttp.send strPayload
    
    ' Get the response
    strResponse = xmlHttp.responseText
    
    ' Parse the JSON response
    Set jsonResponse = JsonConverter.ParseJson(strResponse)
    
    ' Extract the assistant's response content
    CallOpenAI_prompt = jsonResponse("choices")(1)("message")("content")

Cleanup:
    ' Cleanup
    Set xmlHttp = Nothing
    Set jsonResponse = Nothing
    Exit Function

SheetError:
    CallOpenAI_prompt = "Error: Worksheet 'Key_models' not found or cell E3/E5 is invalid."
    Resume Cleanup

ErrorHandler:
    CallOpenAI_prompt = "Error: " & Err.Description
    Resume Cleanup
End Function

