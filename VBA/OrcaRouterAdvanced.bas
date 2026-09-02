Attribute VB_Name = "OrcaRouterAdvanced"
Option Explicit

'===============================================================================
' OrcaRouter Advanced API Learning Sample for Excel VBA
'
' Chat mode remains in OrcaRouterSample.bas.
' This module adds:
'   - Streaming (OpenAI-compatible SSE)
'   - Tool Calling (calculate_sum demo)
'   - Rich HTTP error diagnostics
'
' The same conceptual steps are used in Web / PowerShell / VBA:
'
' STEP 1 - Validate inputs
' STEP 2 - Build request
' STEP 3 - Send HTTP POST
' STEP 4 - Receive response
' STEP 5 - Parse / process result
' STEP 6 - Update UI and trace
'
' Transport note:
' Chat, Streaming, and Tool Calling all use asynchronous MSXML2.XMLHTTP.6.0.
'
' Streaming reads XMLHTTP.responseText incrementally while readyState = 3
' (LOADING). This keeps the sample inside Excel/VBA, avoids an external
' curl.exe process, and lets MSXML decode the UTF-8 HTTP response into a VBA
' Unicode String before text is written to worksheet cells.
'===============================================================================

Private Const API_ENDPOINT_ADV As String = "https://api.orcarouter.ai/v1/chat/completions"
Private Const API_KEY_PLACEHOLDER_ADV As String = "xxx-your-orcarouter-api-key-xxx"
Private Const SAMPLE_SHEET_NAME_ADV As String = "OrcaRouter Chat"
Private Const MAX_STREAM_TRACE_EVENTS As Long = 50
Private Const STREAM_MAX_TIME_SECONDS As Long = 90
Private Const STREAM_WAIT_TRACE_INTERVAL_SECONDS As Long = 15
Private Const STREAM_POLL_INTERVAL_MS As Long = 50
Private Const TOOL_REQUEST_TIMEOUT_SECONDS As Long = 120

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Sub SendOrcaRouterStreaming()

    Dim ws As Worksheet
    Dim httpRequest As Object

    Dim apiKey As String
    Dim model As String
    Dim question As String
    Dim requestBody As String

    Dim currentResponseText As String
    Dim newResponseText As String
    Dim pendingText As String
    Dim lineText As String
    Dim payload As String
    Dim latestJsonPayload As String
    Dim contentPart As String
    Dim answerText As String
    Dim responseHeaders As String
    Dim developerResponse As String
    Dim errorDisplayText As String

    Dim processedChars As Long
    Dim lineFeedPosition As Long
    Dim eventCount As Long
    Dim httpStatus As Long
    Dim readyStateValue As Long

    Dim startedAt As Double
    Dim elapsedSecondsValue As Double
    Dim nextWaitTraceAt As Double
    Dim responseTextAvailable As Boolean

    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME_ADV)

    apiKey = Trim$(CStr(ws.Range("B3").Value))
    model = Trim$(CStr(ws.Range("B4").Value))
    question = Trim$(CStr(ws.Range("B6").Value))

    'Keep committed conversation visible while the request is pending.
    ws.Range("B11").Value = GetOrcaRouterConversationDisplay()
    PrepareRawResponse ws, "Raw JSON - Streaming (latest SSE event)"

    'STEP 1: Validate inputs.
    AddTrace ws, "STEP 1", "LOCAL", "Validate inputs", _
             "API Key: " & MaskApiKey(apiKey) & vbCrLf & _
             "Model: " & model & vbCrLf & _
             "Mode: Streaming" & vbCrLf & _
             "Question length: " & Len(question)

    ValidateAdvancedInputs apiKey, model, question

    'STEP 2: Build request.
    requestBody = BuildStreamingRequestJson(model, question)

    AddTrace ws, "STEP 2", "REQUEST", "Build Streaming request", _
             "Method: POST" & vbCrLf & _
             "Endpoint: " & API_ENDPOINT_ADV & vbCrLf & _
             "Authorization: Bearer " & MaskApiKey(apiKey) & vbCrLf & _
             "Content-Type: application/json; charset=utf-8" & vbCrLf & _
             "Accept: text/event-stream" & vbCrLf & _
             "SSE: data: {...} / data: [DONE]" & vbCrLf & _
             "Body:" & vbCrLf & requestBody

    'STEP 3: Start the asynchronous Streaming request.
    AddTrace ws, "STEP 3", "REQUEST", _
             "Send Streaming POST with MSXML2.XMLHTTP.6.0", _
             "Transport: VBA COM object only" & vbCrLf & _
             "Asynchronous: True" & vbCrLf & _
             "Polling interval: " & STREAM_POLL_INTERVAL_MS & " ms" & vbCrLf & _
             "Overall timeout: " & STREAM_MAX_TIME_SECONDS & " sec" & vbCrLf & _
             "No PowerShell, WScript.Shell, or curl.exe is used."

    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    With httpRequest
        .Open "POST", API_ENDPOINT_ADV, True
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Accept", "text/event-stream"
        .setRequestHeader "X-OrcaRouter-Include-Cost", "true"
        .Send StringToUtf8Bytes(requestBody)
    End With

    startedAt = Timer
    nextWaitTraceAt = STREAM_WAIT_TRACE_INTERVAL_SECONDS

    'STEP 4: Read responseText incrementally.
    '
    'XMLHTTP readyState:
    '  1 = OPENED
    '  2 = HEADERS_RECEIVED
    '  3 = LOADING (partial responseText can be available)
    '  4 = DONE
    '
    'responseText is a VBA Unicode String. Letting MSXML decode the HTTP body
    'avoids the console-code-page mojibake that can occur when curl stdout is
    'read through WScript.Shell.
    Do

        readyStateValue = httpRequest.readyState

        If readyStateValue >= 3 Then

            responseTextAvailable = TryReadXmlHttpResponseText( _
                                        httpRequest, _
                                        currentResponseText)

            If responseTextAvailable Then

                If Len(currentResponseText) > processedChars Then

                    newResponseText = Mid$( _
                                        currentResponseText, _
                                        processedChars + 1)

                    processedChars = Len(currentResponseText)
                    pendingText = pendingText & newResponseText

                    'SSE records are line based. Keep an incomplete final line
                    'in pendingText until the next response fragment arrives.
                    Do

                        lineFeedPosition = InStr(1, pendingText, vbLf, vbBinaryCompare)

                        If lineFeedPosition = 0 Then
                            Exit Do
                        End If

                        lineText = Left$(pendingText, lineFeedPosition - 1)
                        pendingText = Mid$(pendingText, lineFeedPosition + 1)

                        If Right$(lineText, 1) = vbCr Then
                            lineText = Left$(lineText, Len(lineText) - 1)
                        End If

                        ProcessStreamingSseLine _
                            ws, _
                            lineText, _
                            question, _
                            answerText, _
                            latestJsonPayload, _
                            eventCount

                    Loop

                End If

            End If

        End If

        elapsedSecondsValue = ElapsedSeconds(startedAt)

        Application.StatusBar = _
            "Receiving OrcaRouter Streaming response - " & _
            Format$(elapsedSecondsValue, "0") & " sec"

        If elapsedSecondsValue >= nextWaitTraceAt Then

            AddTrace ws, "WAIT", "WAIT", _
                     "Receiving OrcaRouter Streaming response", _
                     "Elapsed: " & Format$(elapsedSecondsValue, "0") & " sec" & vbCrLf & _
                     "readyState: " & CStr(readyStateValue) & vbCrLf & _
                     "SSE events: " & eventCount & vbCrLf & _
                     "Answer chars: " & Len(answerText)

            nextWaitTraceAt = _
                nextWaitTraceAt + STREAM_WAIT_TRACE_INTERVAL_SECONDS

        End If

        If readyStateValue = 4 Then
            Exit Do
        End If

        If elapsedSecondsValue >= STREAM_MAX_TIME_SECONDS Then

            httpRequest.abort
            Application.StatusBar = False

            Err.Raise vbObjectError + 2101, _
                      "SendOrcaRouterStreaming", _
                      "Streaming timed out after " & _
                      STREAM_MAX_TIME_SECONDS & " seconds."

        End If

        Sleep STREAM_POLL_INTERVAL_MS
        YieldToExcel True

    Loop

    Application.StatusBar = False

    'Process an unterminated final line, if one exists.
    If Len(pendingText) > 0 Then

        If Right$(pendingText, 1) = vbCr Then
            pendingText = Left$(pendingText, Len(pendingText) - 1)
        End If

        ProcessStreamingSseLine _
            ws, _
            pendingText, _
            question, _
            answerText, _
            latestJsonPayload, _
            eventCount

    End If

    httpStatus = httpRequest.Status
    responseHeaders = httpRequest.getAllResponseHeaders

    'A non-2xx Streaming response can be ordinary JSON rather than SSE.
    If httpStatus < 200 Or httpStatus >= 300 Then

        currentResponseText = httpRequest.responseText

        DisplayRawResponse ws, _
                           currentResponseText, _
                           "Raw JSON - Streaming error", _
                           httpStatus

        RaiseOrcaRouterHttpError _
            httpStatus, _
            responseHeaders, _
            currentResponseText

    End If

    If Len(latestJsonPayload) > 0 Then

        DisplayRawResponse ws, _
                           latestJsonPayload, _
                           "Raw JSON - Streaming (latest SSE event)", _
                           httpStatus

    Else

        DisplayRawResponse ws, _
                           httpRequest.responseText, _
                           "Raw response - Streaming", _
                           httpStatus

    End If

    AddTrace ws, "STEP 4", "RESPONSE", _
             "Streaming receive completed", _
             "HTTP Status: " & httpStatus & vbCrLf & _
             "Events: " & eventCount & vbCrLf & _
             "Elapsed: " & _
                 Format$(ElapsedSeconds(startedAt), "0.000") & " sec" & vbCrLf & _
             "Headers:" & vbCrLf & responseHeaders

    'STEP 5: Parse / process result.
    AddTrace ws, "STEP 5", "LOCAL", _
             "Aggregate Streaming result", _
             "Answer length: " & Len(answerText)

    If Len(answerText) = 0 Then

        Err.Raise vbObjectError + 2102, _
                  "SendOrcaRouterStreaming", _
                  "Streaming completed but no answer text was captured. " & _
                  "Check Raw JSON and the worksheet trace."

    End If

    'STEP 6: Update UI and trace.
    CommitOrcaRouterConversationTurn question, answerText
    ws.Range("B11").Value = GetOrcaRouterConversationDisplay()

    If Len(latestJsonPayload) > 0 Then
        developerResponse = latestJsonPayload
    Else
        developerResponse = httpRequest.responseText
    End If

    UpdateOrcaRouterDeveloperInformation _
        ws, _
        httpStatus, _
        ElapsedSeconds(startedAt), _
        model, _
        requestBody, _
        developerResponse

    AddTrace ws, "STEP 6", "LOCAL", _
             "Display answer in worksheet", _
             "Mode: Streaming" & vbCrLf & _
             "Completed: True" & vbCrLf & _
             "Total elapsed: " & _
                 Format$(ElapsedSeconds(startedAt), "0.000") & " sec"

CleanExit:
    Application.StatusBar = False
    Set httpRequest = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next

    If Not ws Is Nothing Then

        If Len(answerText) > 0 Then
            errorDisplayText = answerText & vbCrLf & vbCrLf & _
                               "[ERROR]" & vbCrLf & errorDescription
        Else
            errorDisplayText = "ERROR: " & errorDescription
        End If

        ws.Range("B11").Value = _
            GetOrcaRouterConversationDisplay(question, errorDisplayText)

        If httpStatus = 0 And Not httpRequest Is Nothing Then
            httpStatus = httpRequest.Status
        End If

        If Len(currentResponseText) > 0 Then
            developerResponse = currentResponseText
        Else
            developerResponse = latestJsonPayload
        End If

        If startedAt > 0 Then
            elapsedSecondsValue = ElapsedSeconds(startedAt)
        Else
            elapsedSecondsValue = 0
        End If

        UpdateOrcaRouterDeveloperInformation _
            ws, _
            httpStatus, _
            elapsedSecondsValue, _
            model, _
            requestBody, _
            developerResponse

        If Len(developerResponse) > 0 Then
            DisplayRawResponse _
                ws, _
                developerResponse, _
                "Raw JSON - Streaming error", _
                httpStatus
        End If

        AddTrace ws, "ERROR", "ERROR", _
                 "Streaming error", _
                 "Err.Number: " & errorNumber & vbCrLf & _
                 "Err.Source: " & errorSource & vbCrLf & _
                 "Err.Description: " & errorDescription & vbCrLf & _
                 "HTTP Status: " & _
                     IIf(httpStatus = 0, "(not available)", CStr(httpStatus))

    End If

    MsgBox "An error occurred during Streaming." & vbCrLf & _
           errorDescription & vbCrLf & vbCrLf & _
           "Check Raw JSON and the worksheet trace for details.", _
           vbExclamation, _
           "OrcaRouter Streaming"

    On Error GoTo 0
    GoTo CleanExit

End Sub

Private Sub ProcessStreamingSseLine( _
    ByVal ws As Worksheet, _
    ByVal lineText As String, _
    ByVal question As String, _
    ByRef answerText As String, _
    ByRef latestJsonPayload As String, _
    ByRef eventCount As Long)

    Dim payload As String
    Dim contentPart As String

    If Left$(lineText, 5) <> "data:" Then
        Exit Sub
    End If

    payload = Trim$(Mid$(lineText, 6))

    If Len(payload) = 0 Then
        Exit Sub
    End If

    If payload = "[DONE]" Then

        AddTrace ws, "STEP 4", "STREAM", _
                 "SSE finished [DONE]", _
                 "Events: " & eventCount

        Exit Sub

    End If

    latestJsonPayload = payload
    eventCount = eventCount + 1

    If eventCount <= MAX_STREAM_TRACE_EVENTS Then

        AddTrace ws, "STEP 4", "STREAM", _
                 "SSE data #" & eventCount, _
                 payload

    ElseIf eventCount = MAX_STREAM_TRACE_EVENTS + 1 Then

        AddTrace ws, "STEP 4", "STREAM", _
                 "Omit additional SSE trace entries", _
                 "For readability, only the first " & _
                 MAX_STREAM_TRACE_EVENTS & _
                 " SSE events are shown individually."

    End If

    If InStr(1, payload, """error""", vbTextCompare) > 0 Then
        RaiseStreamingError payload
    End If

    If TryExtractJsonStringProperty(payload, "content", contentPart) Then

        If Len(contentPart) > 0 Then

            answerText = answerText & contentPart
            ws.Range("B11").Value = _
                GetOrcaRouterConversationDisplay(question, answerText)

            YieldToExcel

        End If

    End If

End Sub

Private Function TryReadXmlHttpResponseText( _
    ByVal httpRequest As Object, _
    ByRef responseText As String) As Boolean

    On Error Resume Next

    Err.Clear
    responseText = httpRequest.responseText

    If Err.Number = 0 Then
        TryReadXmlHttpResponseText = True
    Else
        responseText = vbNullString
        TryReadXmlHttpResponseText = False
    End If

    Err.Clear
    On Error GoTo 0

End Function

Public Sub SendOrcaRouterToolCalling()

    Dim ws As Worksheet
    Dim httpRequest As Object

    Dim apiKey As String
    Dim model As String
    Dim question As String

    Dim firstRequest As String
    Dim firstResponse As String
    Dim firstHeaders As String
    Dim firstStatus As Long

    Dim toolSection As String
    Dim toolCallId As String
    Dim functionName As String
    Dim argumentsJson As String

    Dim valueA As Double
    Dim valueB As Double
    Dim sumValue As Double
    Dim toolResultJson As String

    Dim secondRequest As String
    Dim secondResponse As String
    Dim secondHeaders As String
    Dim secondStatus As Long
    Dim assistantText As String
    Dim developerRequest As String
    Dim developerResponse As String
    Dim developerStatus As Long
    Dim elapsedSecondsValue As Double

    Dim startedAt As Double

    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME_ADV)

    apiKey = Trim$(CStr(ws.Range("B3").Value))
    model = Trim$(CStr(ws.Range("B4").Value))
    question = Trim$(CStr(ws.Range("B6").Value))

    'Keep committed conversation visible while the request is pending.
    ws.Range("B11").Value = GetOrcaRouterConversationDisplay()
    PrepareRawResponse ws, "Raw JSON - Tool Calling"

    'STEP 1: Validate inputs.
    AddTrace ws, "STEP 1", "LOCAL", "Validate inputs", _
             "API Key: " & MaskApiKey(apiKey) & vbCrLf & _
             "Model: " & model & vbCrLf & _
             "Mode: Tool Calling" & vbCrLf & _
             "Question length: " & Len(question)

    ValidateAdvancedInputs apiKey, model, question

    startedAt = Timer

    'STEP 2: Build request.
    firstRequest = BuildToolRequestJson(model, question)

    AddTrace ws, "STEP 2", "REQUEST", _
             "Build Tool Calling request #1", _
             "Endpoint: " & API_ENDPOINT_ADV & vbCrLf & _
             "Tool: calculate_sum(a, b)" & vbCrLf & _
             "tool_choice: calculate_sum" & vbCrLf & _
             "Body:" & vbCrLf & firstRequest

    'STEP 3 / STEP 4: First HTTP request.
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    SendJsonRequest httpRequest, apiKey, firstRequest, _
                    firstStatus, firstHeaders, firstResponse

    DisplayRawResponse ws, firstResponse, _
                       "Raw JSON - Tool Calling response #1", _
                       firstStatus

    AddTrace ws, "STEP 4", "RESPONSE", _
             "Receive Tool Calling response #1", _
             "HTTP Status: " & firstStatus & vbCrLf & _
             "Headers:" & vbCrLf & firstHeaders & vbCrLf & _
             "Raw response:" & vbCrLf & firstResponse

    If firstStatus < 200 Or firstStatus >= 300 Then
        RaiseOrcaRouterHttpError firstStatus, firstHeaders, firstResponse
    End If

    'STEP 5A: Parse tool call.
    toolSection = GetToolCallsSection(firstResponse)

    If Not TryExtractJsonStringProperty(toolSection, "id", toolCallId) Then
        Err.Raise vbObjectError + 2201, "SendOrcaRouterToolCalling", _
                  "Could not extract tool_call id."
    End If

    If Not TryExtractJsonStringProperty(toolSection, "name", functionName) Then
        Err.Raise vbObjectError + 2202, "SendOrcaRouterToolCalling", _
                  "Could not extract tool function name."
    End If

    If Not TryExtractJsonStringProperty(toolSection, "arguments", argumentsJson) Then
        Err.Raise vbObjectError + 2203, "SendOrcaRouterToolCalling", _
                  "Could not extract tool arguments."
    End If

    AddTrace ws, "STEP 5A", "TOOL", _
             "Receive Tool Call from model", _
             "Tool Call ID: " & toolCallId & vbCrLf & _
             "Function: " & functionName & vbCrLf & _
             "Arguments: " & argumentsJson

    If StrComp(functionName, "calculate_sum", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2204, "SendOrcaRouterToolCalling", _
                  "Unsupported Tool requested: " & functionName
    End If

    valueA = ExtractJsonNumber(argumentsJson, "a")
    valueB = ExtractJsonNumber(argumentsJson, "b")
    sumValue = valueA + valueB

    toolResultJson = _
        "{""a"":" & JsonNumber(valueA) & _
        ",""b"":" & JsonNumber(valueB) & _
        ",""sum"":" & JsonNumber(sumValue) & "}"

    'STEP 5B: Execute local tool.
    AddTrace ws, "STEP 5B", "TOOL", _
             "Execute local calculate_sum function", _
             "a = " & valueA & vbCrLf & _
             "b = " & valueB & vbCrLf & _
             "sum = " & sumValue & vbCrLf & _
             "Tool result JSON: " & toolResultJson

    'STEP 5C: Build and send the second request with tool result.
    secondRequest = BuildToolResultRequestJson( _
                        model, _
                        question, _
                        toolCallId, _
                        functionName, _
                        argumentsJson, _
                        toolResultJson)

    AddTrace ws, "STEP 5C", "REQUEST", _
             "Build request #2 with Tool result", _
             "Body:" & vbCrLf & secondRequest

    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    SendJsonRequest httpRequest, apiKey, secondRequest, _
                    secondStatus, secondHeaders, secondResponse

    DisplayRawResponse ws, secondResponse, _
                       "Raw JSON - Tool Calling response #2 (final)", _
                       secondStatus

    AddTrace ws, "STEP 4", "RESPONSE", _
             "Receive Tool Calling response #2", _
             "HTTP Status: " & secondStatus & vbCrLf & _
             "Headers:" & vbCrLf & secondHeaders & vbCrLf & _
             "Raw response:" & vbCrLf & secondResponse

    If secondStatus < 200 Or secondStatus >= 300 Then
        RaiseOrcaRouterHttpError secondStatus, secondHeaders, secondResponse
    End If

    'STEP 5: Parse final assistant answer.
    assistantText = ExtractAssistantContent(secondResponse)

    AddTrace ws, "STEP 5", "LOCAL", _
             "Parse final answer after Tool Calling", _
             "Answer length: " & Len(assistantText)

    'STEP 6: Update UI and trace.
    CommitOrcaRouterConversationTurn question, assistantText
    ws.Range("B11").Value = GetOrcaRouterConversationDisplay()

    developerRequest = _
        "{""request_1"":" & firstRequest & _
        ",""request_2"":" & secondRequest & "}"

    UpdateOrcaRouterDeveloperInformation _
        ws, _
        secondStatus, _
        ElapsedSeconds(startedAt), _
        model, _
        developerRequest, _
        secondResponse, _
        firstResponse

    AddTrace ws, "STEP 6", "LOCAL", "Display answer in worksheet", _
             "Mode: Tool Calling" & vbCrLf & _
             "Completed: True" & vbCrLf & _
             "Total elapsed: " & Format$(ElapsedSeconds(startedAt), "0.000") & " sec"

CleanExit:
    Application.StatusBar = False
    Set httpRequest = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next

    If Not ws Is Nothing Then

        ws.Range("B11").Value = _
            GetOrcaRouterConversationDisplay( _
                question, _
                "ERROR: " & errorDescription)

        If secondStatus > 0 Then
            developerStatus = secondStatus
            developerResponse = secondResponse
        Else
            developerStatus = firstStatus
            developerResponse = firstResponse
        End If

        If Len(secondRequest) > 0 Then
            developerRequest = _
                "{""request_1"":" & firstRequest & _
                ",""request_2"":" & secondRequest & "}"
        Else
            developerRequest = firstRequest
        End If

        If startedAt > 0 Then
            elapsedSecondsValue = ElapsedSeconds(startedAt)
        Else
            elapsedSecondsValue = 0
        End If

        UpdateOrcaRouterDeveloperInformation _
            ws, _
            developerStatus, _
            elapsedSecondsValue, _
            model, _
            developerRequest, _
            developerResponse

        AddTrace ws, "ERROR", "ERROR", "Tool Calling error", _
                 "Err.Number: " & errorNumber & vbCrLf & _
                 "Err.Source: " & errorSource & vbCrLf & _
                 "Err.Description: " & errorDescription & vbCrLf & _
                 "First HTTP Status: " & firstStatus & vbCrLf & _
                 "Second HTTP Status: " & secondStatus

    End If

    MsgBox "An error occurred during Tool Calling." & vbCrLf & _
           errorDescription & vbCrLf & vbCrLf & _
           "Check the worksheet trace for details.", _
           vbExclamation, _
           "OrcaRouter Tool Calling"

    On Error GoTo 0
    GoTo CleanExit

End Sub

Public Sub RunOrcaRouterAdvancedSelfTests()

    Dim streamingJson As String
    Dim toolJson As String
    Dim toolResultJson As String
    Dim extractedText As String
    Dim extractedNumber As Double
    Dim headerValue As String

    'Verify the common Answer extraction shapes used by Chat/Tool Calling.
    If ExtractAssistantContent( _
           "{""choices"":[{""message"":{""role"":""assistant"",""content"":""hello""}}]}") <> "hello" Then

        Err.Raise vbObjectError + 3110, "RunOrcaRouterAdvancedSelfTests", _
                  "ExtractAssistantContent failed for string content."
    End If

    If ExtractAssistantContent( _
           "{""choices"":[{""message"":{""role"":""assistant"",""content"":[{""type"":""text"",""text"":""hello""},{""type"":""text"",""text"":""array""}]}}]}") <> _
       "hello" & vbCrLf & "array" Then

        Err.Raise vbObjectError + 3111, "RunOrcaRouterAdvancedSelfTests", _
                  "ExtractAssistantContent failed for multi-part array content."
    End If

    streamingJson = BuildStreamingRequestJson("orcarouter/free", "hello")

    If InStr(1, streamingJson, """stream"":true", vbBinaryCompare) = 0 Or _
       InStr(1, streamingJson, """include_usage"":true", vbBinaryCompare) = 0 Then

        Err.Raise vbObjectError + 3101, "RunOrcaRouterAdvancedSelfTests", _
                  "BuildStreamingRequestJson produced an unexpected result."
    End If

    toolJson = BuildToolRequestJson("orcarouter/free", "add 1 and 2")

    If InStr(1, toolJson, """name"":""calculate_sum""", vbBinaryCompare) = 0 Or _
       InStr(1, toolJson, """tool_choice""", vbBinaryCompare) = 0 Then

        Err.Raise vbObjectError + 3102, "RunOrcaRouterAdvancedSelfTests", _
                  "BuildToolRequestJson produced an unexpected result."
    End If

    toolResultJson = BuildToolResultRequestJson( _
                         "orcarouter/free", _
                         "add 1 and 2", _
                         "call_123", _
                         "calculate_sum", _
                         "{""a"":1,""b"":2}", _
                         "{""a"":1,""b"":2,""sum"":3}")

    If InStr(1, toolResultJson, """tool_call_id"":""call_123""", vbBinaryCompare) = 0 Or _
       InStr(1, toolResultJson, """role"":""tool""", vbBinaryCompare) = 0 Then

        Err.Raise vbObjectError + 3103, "RunOrcaRouterAdvancedSelfTests", _
                  "BuildToolResultRequestJson produced an unexpected result."
    End If

    If Not TryExtractJsonStringProperty( _
               "{""message"":""hello""}", _
               "message", _
               extractedText) Then

        Err.Raise vbObjectError + 3104, "RunOrcaRouterAdvancedSelfTests", _
                  "TryExtractJsonStringProperty did not find a known property."
    End If

    If extractedText <> "hello" Then
        Err.Raise vbObjectError + 3105, "RunOrcaRouterAdvancedSelfTests", _
                  "TryExtractJsonStringProperty decoded an unexpected value."
    End If

    extractedNumber = ExtractJsonNumber("{""a"":123.5}", "a")

    If Abs(extractedNumber - 123.5) > 0.000001 Then
        Err.Raise vbObjectError + 3106, "RunOrcaRouterAdvancedSelfTests", _
                  "ExtractJsonNumber produced an unexpected value."
    End If

    headerValue = ExtractHeaderValue( _
                      "Content-Type: application/json" & vbCrLf & _
                      "Retry-After: 10" & vbCrLf, _
                      "Retry-After")

    If headerValue <> "10" Then
        Err.Raise vbObjectError + 3108, "RunOrcaRouterAdvancedSelfTests", _
                  "ExtractHeaderValue produced an unexpected value."
    End If

    If JsonNumber(123.5) <> "123.5" Then
        Err.Raise vbObjectError + 3109, "RunOrcaRouterAdvancedSelfTests", _
                  "JsonNumber did not produce locale-independent JSON."
    End If

End Sub

Private Sub ValidateAdvancedInputs( _
    ByVal apiKey As String, _
    ByVal model As String, _
    ByVal question As String)

    If Len(apiKey) = 0 Or _
       apiKey = API_KEY_PLACEHOLDER_ADV Or _
       Left$(apiKey, 4) = "xxx-" Then

        Err.Raise vbObjectError + 2001, "ValidateAdvancedInputs", _
                  "The API key is still the dummy value. Enter your OrcaRouter API key in cell B3."

    End If

    If Len(model) = 0 Then
        Err.Raise vbObjectError + 2002, "ValidateAdvancedInputs", _
                  "Model is empty. Enter a model name in cell B4."
    End If

    If Len(question) = 0 Then
        Err.Raise vbObjectError + 2003, "ValidateAdvancedInputs", _
                  "Question is empty. Enter a question in cell B6."
    End If

End Sub

Private Function BuildStreamingRequestJson( _
    ByVal model As String, _
    ByVal question As String) As String

    Dim result As String

    result = "{"
    result = result & """model"":""" & JsonEscape(model) & ""","
    result = result & """messages"":["
    result = result & BuildOrcaRouterConversationMessageItemsJson(question)
    result = result & "],"
    result = result & """stream"":true,"
    result = result & """stream_options"":{""include_usage"":true}"
    result = result & "}"

    BuildStreamingRequestJson = result

End Function

Private Function BuildToolRequestJson( _
    ByVal model As String, _
    ByVal question As String) As String

    Dim result As String

    'Keep each VBA statement short. VBA allows at most 24 line-continuation
    'characters in a single logical statement.
    result = "{"
    result = result & """model"":""" & JsonEscape(model) & ""","
    result = result & """messages"":["
    result = result & "{"
    result = result & """role"":""system"","
    result = result & """content"":""This is a Tool Calling learning demo. Call calculate_sum. If the user did not provide two numbers, use 123 and 456."""
    result = result & "},"
    result = result & BuildOrcaRouterConversationMessageItemsJson(question)
    result = result & "],"
    result = result & """tools"":[{"
    result = result & """type"":""function"","
    result = result & """function"":{"
    result = result & """name"":""calculate_sum"","
    result = result & """description"":""Add two numbers and return the sum."","
    result = result & """parameters"":{"
    result = result & """type"":""object"","
    result = result & """properties"":{"
    result = result & """a"":{""type"":""number"",""description"":""First number""},"
    result = result & """b"":{""type"":""number"",""description"":""Second number""}"
    result = result & "},"
    result = result & """required"":[""a"",""b""],"
    result = result & """additionalProperties"":false"
    result = result & "}"
    result = result & "}"
    result = result & "}],"
    result = result & """tool_choice"":{"
    result = result & """type"":""function"","
    result = result & """function"":{""name"":""calculate_sum""}"
    result = result & "}"
    result = result & "}"

    BuildToolRequestJson = result

End Function

Private Function BuildToolResultRequestJson( _
    ByVal model As String, _
    ByVal question As String, _
    ByVal toolCallId As String, _
    ByVal functionName As String, _
    ByVal argumentsJson As String, _
    ByVal toolResultJson As String) As String

    Dim result As String

    'Build the JSON in multiple statements so the source stays within
    'VBA's line-continuation limit.
    result = "{"
    result = result & """model"":""" & JsonEscape(model) & ""","
    result = result & """messages"":["
    result = result & "{"
    result = result & """role"":""system"","
    result = result & """content"":""This is a Tool Calling learning demo. Call calculate_sum. If the user did not provide two numbers, use 123 and 456."""
    result = result & "},"
    result = result & BuildOrcaRouterConversationMessageItemsJson(question)
    result = result & ","
    result = result & "{"
    result = result & """role"":""assistant"","
    result = result & """content"":null,"
    result = result & """tool_calls"":[{"
    result = result & """id"":""" & JsonEscape(toolCallId) & ""","
    result = result & """type"":""function"","
    result = result & """function"":{"
    result = result & """name"":""" & JsonEscape(functionName) & ""","
    result = result & """arguments"":""" & JsonEscape(argumentsJson) & """"
    result = result & "}"
    result = result & "}]"
    result = result & "},"
    result = result & "{"
    result = result & """role"":""tool"","
    result = result & """tool_call_id"":""" & JsonEscape(toolCallId) & ""","
    result = result & """content"":""" & JsonEscape(toolResultJson) & """"
    result = result & "}"
    result = result & "]"
    result = result & "}"

    BuildToolResultRequestJson = result

End Function

Private Sub SendJsonRequest( _
    ByVal httpRequest As Object, _
    ByVal apiKey As String, _
    ByVal requestBody As String, _
    ByRef httpStatus As Long, _
    ByRef responseHeaders As String, _
    ByRef responseText As String)

    Dim ws As Worksheet
    Dim startedAt As Double

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME_ADV)

    AddTrace ws, _
             "STEP 3", "REQUEST", "Send XMLHTTP POST", _
             "Endpoint: " & API_ENDPOINT_ADV & vbCrLf & _
             "Authorization: Bearer " & MaskApiKey(apiKey) & vbCrLf & _
             "Transport: MSXML2.XMLHTTP.6.0" & vbCrLf & _
             "Excel remains responsive while waiting."

    startedAt = Timer

    With httpRequest

        'Use the same asynchronous XMLHTTP transport as normal Chat.
        '
        'Tool Calling can make two HTTP requests. A synchronous Send on either
        'request would freeze the Excel UI until that call completes.
        '
        'With asynchronous=True, Send returns to VBA immediately. The shared
        'WaitForXmlHttpResponse helper polls readyState, updates Trace and the
        'StatusBar, and yields to Excel between polls.
        .Open "POST", API_ENDPOINT_ADV, True
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "X-OrcaRouter-Include-Cost", "true"
        .Send StringToUtf8Bytes(requestBody)

    End With

    WaitForXmlHttpResponse _
        httpRequest, _
        ws, _
        "Waiting for OrcaRouter Tool Calling response", _
        startedAt, _
        TOOL_REQUEST_TIMEOUT_SECONDS, _
        15

    httpStatus = httpRequest.Status
    responseHeaders = httpRequest.getAllResponseHeaders
    responseText = httpRequest.responseText

    Set ws = Nothing

End Sub

Public Sub RaiseOrcaRouterHttpError( _
    ByVal httpStatus As Long, _
    ByVal responseHeaders As String, _
    ByVal responseText As String)

    Dim errorMessage As String
    Dim errorType As String
    Dim errorCode As String
    Dim retryAfter As String
    Dim guidance As String

    Call TryExtractJsonStringProperty(responseText, "message", errorMessage)
    Call TryExtractJsonStringProperty(responseText, "type", errorType)
    Call TryExtractJsonStringProperty(responseText, "code", errorCode)

    retryAfter = ExtractHeaderValue(responseHeaders, "Retry-After")
    guidance = BuildHttpGuidance(httpStatus, errorCode, retryAfter, errorMessage)

    Err.Raise vbObjectError + 2300 + httpStatus, "RaiseOrcaRouterHttpError", _
              "HTTP " & httpStatus & vbCrLf & _
              "error.type: " & IIf(Len(errorType) = 0, "(empty)", errorType) & vbCrLf & _
              "error.code: " & IIf(Len(errorCode) = 0, "(empty)", errorCode) & vbCrLf & _
              "message: " & IIf(Len(errorMessage) = 0, responseText, errorMessage) & vbCrLf & _
              "Retry-After: " & IIf(Len(retryAfter) = 0, "(none)", retryAfter) & vbCrLf & _
              "Guidance: " & guidance

End Sub

Private Function BuildHttpGuidance( _
    ByVal httpStatus As Long, _
    ByVal errorCode As String, _
    ByVal retryAfter As String, _
    ByVal errorMessage As String) As String

    'Prefer the specific OrcaRouter error.code over the HTTP status.
    'The live API can occasionally surface quota conditions with a status that
    'differs from older documentation. Matching the code keeps the guidance
    'useful for both 402/403-style quota responses.
    If StrComp(errorCode, "free_quota_exhausted", vbTextCompare) = 0 Then

        BuildHttpGuidance = _
            "The orcarouter/free allowance/capacity is unavailable for this request. " & _
            "The Tool Calling code did not run yet because OrcaRouter rejected " & _
            "the request before model execution. Do not retry in a loop. " & _
            "Use a specific paid chat-capable model only after you intentionally " & _
            "add wallet credit, or wait until free access is available again."

        Exit Function

    End If

    Select Case httpStatus

        Case 400

            Select Case errorCode

                Case "bad_request_body"
                    BuildHttpGuidance = _
                        "The request JSON could not be parsed. Check the request body in the trace."

                Case "model_price_error"
                    BuildHttpGuidance = _
                        "Pricing is not configured for this model. Contact OrcaRouter support."

                Case "api_not_implemented"
                    BuildHttpGuidance = _
                        "This endpoint or operation is not supported for the selected model."

                Case "prompt_blocked", "sensitive_words_detected", "guardrail_blocked"
                    BuildHttpGuidance = _
                        "The request was blocked by a provider safety policy or workspace guardrail. Change the input or policy."

                Case "firewall_blocked"
                    BuildHttpGuidance = _
                        "The Agent Firewall denied the requested tool. Review the firewall policy and error metadata."

                Case "firewall_approval_pending"
                    BuildHttpGuidance = _
                        "The tool call is waiting for firewall approval. A plain retry will not resolve it."

                Case Else
                    BuildHttpGuidance = _
                        "Bad request. Check error.code, message, and the request JSON in the trace."

            End Select

        Case 401
            BuildHttpGuidance = _
                "The API key is invalid or the Authorization header is incorrect."

        Case 402
            BuildHttpGuidance = _
                "Payment or quota is required for this request. Check error.code and " & _
                "the OrcaRouter billing/quota message before changing models."

        Case 403

            If InStr(1, errorMessage, "token cycle spend limit reached", vbTextCompare) > 0 Then

                BuildHttpGuidance = _
                    "This API key reached its recurring spend limit. Wait for the reset time in the message or raise the key limit."

            Else

                Select Case errorCode

                    Case "insufficient_user_quota"
                        BuildHttpGuidance = _
                            "Check workspace balance and member/agent budget."

                    Case "pre_consume_token_quota_failed"
                        BuildHttpGuidance = _
                            "Check the quota limit assigned to this API key."

                    Case "access_denied"
                        BuildHttpGuidance = _
                            "The key is valid but this request is not permitted. Check cycle limits, IP allowlist, and model access."

                    Case Else
                        BuildHttpGuidance = _
                            "HTTP 403 has multiple possible causes. Check model access, error.code, and message."

                End Select

            End If

        Case 404
            BuildHttpGuidance = _
                "The endpoint or model was not found. Verify the endpoint and model id."

        Case 425
            BuildHttpGuidance = _
                "The selected model is announced but not live yet. Check error metadata for an alternative."

        Case 429

            If Len(retryAfter) > 0 Then
                BuildHttpGuidance = _
                    "Rate limit reached. Wait for Retry-After seconds, then retry once."
            Else
                BuildHttpGuidance = _
                    "A free-tier 429 without Retry-After is not time-based. Shorten the prompt before retrying."
            End If

        Case 500
            BuildHttpGuidance = _
                "OrcaRouter returned an internal server error."

        Case 502
            BuildHttpGuidance = _
                "All upstream providers or fallback routes failed. Retry later or inspect fallback headers."

        Case 503

            Select Case errorCode

                Case "model_not_found"
                    BuildHttpGuidance = _
                        "Check whether the selected model is available for the current account."

                Case "byok:key_unavailable"
                    BuildHttpGuidance = _
                        "The workspace BYOK provider key could not be used. Rotate or re-add the provider key, or review platform fallback settings."

                Case Else
                    BuildHttpGuidance = _
                        "The service or upstream provider may be temporarily unavailable."

            End Select

        Case Else
            BuildHttpGuidance = _
                "Check error.code, error.type, and HTTP Status in the trace."

    End Select

End Function

Private Sub RaiseStreamingError(ByVal payload As String)

    Dim errorMessage As String
    Dim errorType As String
    Dim errorCode As String

    Call TryExtractJsonStringProperty(payload, "message", errorMessage)
    Call TryExtractJsonStringProperty(payload, "type", errorType)
    Call TryExtractJsonStringProperty(payload, "code", errorCode)

    Err.Raise vbObjectError + 2401, "RaiseStreamingError", _
              "Streaming error" & vbCrLf & _
              "type: " & errorType & vbCrLf & _
              "code: " & errorCode & vbCrLf & _
              "message: " & errorMessage

End Sub

Private Function GetToolCallsSection(ByVal responseJson As String) As String

    Dim startPosition As Long

    startPosition = InStr(1, responseJson, """tool_calls""", vbTextCompare)

    If startPosition = 0 Then
        Err.Raise vbObjectError + 2501, "GetToolCallsSection", _
                  "tool_calls was not returned. Check whether the selected model supports Tool Calling."
    End If

    GetToolCallsSection = Mid$(responseJson, startPosition)

End Function

Private Function TryExtractJsonStringProperty( _
    ByVal jsonText As String, _
    ByVal propertyName As String, _
    ByRef decodedValue As String) As Boolean

    Dim regularExpression As Object
    Dim matches As Object
    Dim pattern As String

    Set regularExpression = CreateObject("VBScript.RegExp")

    pattern = _
        Chr$(34) & propertyName & Chr$(34) & _
        Chr$(92) & "s*:" & Chr$(92) & "s*" & _
        Chr$(34) & _
        "((" & Chr$(92) & Chr$(92) & ".|[^" & _
        Chr$(34) & Chr$(92) & Chr$(92) & "])*)" & _
        Chr$(34)

    With regularExpression
        .Global = False
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = pattern
    End With

    Set matches = regularExpression.Execute(jsonText)

    If matches.Count = 0 Then

        decodedValue = vbNullString
        TryExtractJsonStringProperty = False

    Else

        decodedValue = JsonUnescape(matches(0).SubMatches(0))
        TryExtractJsonStringProperty = True

    End If

    Set matches = Nothing
    Set regularExpression = Nothing

End Function

Private Function ExtractJsonNumber( _
    ByVal jsonText As String, _
    ByVal propertyName As String) As Double

    Dim regularExpression As Object
    Dim matches As Object
    Dim pattern As String
    Dim numberText As String

    Set regularExpression = CreateObject("VBScript.RegExp")

    pattern = _
        Chr$(34) & propertyName & Chr$(34) & _
        Chr$(92) & "s*:" & Chr$(92) & "s*" & _
        "(-?[0-9]+(" & Chr$(92) & ".[0-9]+)?([eE][+-]?[0-9]+)?)"

    With regularExpression
        .Global = False
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = pattern
    End With

    Set matches = regularExpression.Execute(jsonText)

    If matches.Count = 0 Then
        Err.Raise vbObjectError + 2601, "ExtractJsonNumber", _
                  "Could not extract numeric Tool argument: " & propertyName
    End If

    numberText = matches(0).SubMatches(0)

    If Application.International(xlDecimalSeparator) <> "." Then
        numberText = Replace$( _
                        numberText, _
                        ".", _
                        Application.International(xlDecimalSeparator))
    End If

    ExtractJsonNumber = CDbl(numberText)

    Set matches = Nothing
    Set regularExpression = Nothing

End Function

Private Function JsonNumber(ByVal value As Double) As String

    Dim valueText As String

    valueText = CStr(value)

    If Application.International(xlDecimalSeparator) <> "." Then
        valueText = Replace$( _
                        valueText, _
                        Application.International(xlDecimalSeparator), _
                        ".")
    End If

    JsonNumber = valueText

End Function

Private Function ExtractHeaderValue( _
    ByVal headersText As String, _
    ByVal headerName As String) As String

    Dim lines() As String
    Dim oneLine As Variant
    Dim colonPosition As Long
    Dim namePart As String

    lines = Split(Replace$(headersText, vbCrLf, vbLf), vbLf)

    For Each oneLine In lines

        colonPosition = InStr(1, CStr(oneLine), ":", vbBinaryCompare)

        If colonPosition > 0 Then

            namePart = Trim$(Left$(CStr(oneLine), colonPosition - 1))

            If StrComp(namePart, headerName, vbTextCompare) = 0 Then
                ExtractHeaderValue = _
                    Trim$(Mid$(CStr(oneLine), colonPosition + 1))
                Exit Function
            End If

        End If

    Next oneLine

End Function
