Function CallOpenAI(ask As String, value As String) As String
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Dim currentCell As Range
    'Dim leftCell As Range
    'Dim leftCellValue As String
    
    'Set currentCell = ActiveCell ' Or any specific cell like Range("B2")
    'Set leftCell = currentCell.Offset(0, -1)
    'leftCellValue = leftCell.Value
    
    Dim url As String
    url = "https://api.openai.com/v1/chat/completions"
    
    ' Prepare the JSON string as per the API's requirements
    Dim jsonBody As String
    jsonBody = "{""model"": ""gpt-4o-mini"", ""messages"": [" & _
        "{""role"": ""user"", ""content"": """ & ask & ": " & value & """}" & _
    "]}"
    
    ' Open and configure the request
    httpRequest.Open "POST", url, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.setRequestHeader "Authorization", "Bearer sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    
    ' Send the request with the JSON body
    httpRequest.Send (jsonBody)
    
    ' Check the status code and process the response
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText
        CallOpenAI = ExtractContentFromJSON(response)
    Else
        MsgBox "Error: " & httpRequest.Status & " - " & httpRequest.statusText
    End If
End Function


Function ExtractContentFromJSON(jsonResponse As String)
    Dim contentStart As Long
    Dim contentEnd As Long
    Dim content As String
    
    ' Find the position of "content": " and adjust to get the start of the actual content
    contentStart = InStr(jsonResponse, """content"": """) + Len("""content"": """)
    
    If contentStart > Len("""content"": """) Then
        ' Find the end of the content based on the structure of the JSON
        contentEnd = InStr(contentStart, jsonResponse, """", vbTextCompare)
        
        ' Extract the content
        content = Mid(jsonResponse, contentStart, contentEnd - contentStart)
        
        ' Output the result
        ExtractContentFromJSON = content
    Else
        ExtractContentFromJSON = ""
        MsgBox "Content not found."
    End If
End Function

