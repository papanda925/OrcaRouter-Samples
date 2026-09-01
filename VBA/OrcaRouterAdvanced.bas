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
' Streaming note:
' WinHttp.WinHttpRequest is intentionally kept for ordinary requests.
' For a dependency-light VBA streaming demo, Windows curl.exe is started and
' its SSE stdout is read line-by-line. The API key is passed through curl's
' stdin config, not on the process command line.
'===============================================================================

Private Const API_ENDPOINT_ADV As String = "https://api.orcarouter.ai/v1/chat/completions"
Private Const API_KEY_PLACEHOLDER_ADV As String = "xxx-your-orcarouter-api-key-xxx"
Private Const SAMPLE_SHEET_NAME_ADV As String = "OrcaRouter Chat"
Private Const MAX_STREAM_TRACE_EVENTS As Long = 50

Public Sub SendOrcaRouterStreaming()

    Dim ws As Worksheet
    Dim apiKey As String
    Dim model As String
    Dim question As String
    Dim requestBody As String

    Dim curlPath As String
    Dim command As String
    Dim configText As String
    Dim shell As Object
    Dim exec As Object

    Dim lineText As String
    Dim payload As String
    Dim contentPart As String
    Dim answerText As String
    Dim stderrText As String
    Dim responseHeaders As String
    Dim nonSseOutput As String
    Dim httpStatus As Long
    Dim readingHeaders As Boolean

    Dim eventCount As Long
    Dim startedAt As Double

    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME_ADV)

    apiKey = Trim$(CStr(ws.Range("B3").Value))
    model = Trim$(CStr(ws.Range("B4").Value))
    question = Trim$(CStr(ws.Range("B6").Value))

    ws.Range("B11").Value = vbNullString

    'STEP 1: Validate inputs.
    AddTrace ws, "STEP 1", "LOCAL", "Validate inputs", _
             "API Key: " & MaskApiKey(apiKey) & vbCrLf & _
             "Model: " & model & vbCrLf & _
             "Mode: Streaming" & vbCrLf & _
             "Question length: " & Len(question)

    ValidateAdvancedInputs apiKey, model, question

    curlPath = FindCurlPath()

    If Len(curlPath) = 0 Then
        Err.Raise vbObjectError + 2101, "SendOrcaRouterStreaming", _
                  "curl.exe was not found. This sample expects the curl.exe included with Windows 10/11."
    End If

    'STEP 2: Build request.
    requestBody = BuildStreamingRequestJson(model, question)

    AddTrace ws, "STEP 2", "REQUEST", "Build Streaming request", _
             "Method: POST" & vbCrLf & _
             "Endpoint: " & API_ENDPOINT_ADV & vbCrLf & _
             "Authorization: Bearer " & MaskApiKey(apiKey) & vbCrLf & _
             "Content-Type: application/json" & vbCrLf & _
             "SSE: data: {...} / data: [DONE]" & vbCrLf & _
             "Body:" & vbCrLf & requestBody

    'STEP 3: Send HTTP POST.
    AddTrace ws, "STEP 3", "REQUEST", "Send Streaming POST with curl.exe", _
             "curl.exe --config - --no-buffer --silent --show-error --include" & vbCrLf & _
             "The API key is passed through curl stdin config, not as a command-line argument."

    command = """" & curlPath & """ --config - --no-buffer --silent --show-error --include"

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(command)

    configText = BuildCurlConfig(apiKey, requestBody)

    exec.StdIn.Write configText
    exec.StdIn.Close

    startedAt = Timer

    'STEP 4: Receive response.
    Do

        If Not exec.StdOut.AtEndOfStream Then

            lineText = exec.StdOut.ReadLine

            If Left$(lineText, 5) = "HTTP/" Then

                httpStatus = ParseHttpStatusLine(lineText)
                responseHeaders = responseHeaders & lineText & vbCrLf
                readingHeaders = True

            ElseIf Len(lineText) = 0 Then

                readingHeaders = False

            ElseIf Left$(lineText, 5) = "data:" Then

                payload = Trim$(Mid$(lineText, 6))

                If payload = "[DONE]" Then

                    AddTrace ws, "STEP 4", "STREAM", "SSE finished [DONE]", _
                             "Events: " & eventCount

                ElseIf Len(payload) > 0 Then

                    eventCount = eventCount + 1

                    If eventCount <= MAX_STREAM_TRACE_EVENTS Then

                        AddTrace ws, "STEP 4", "STREAM", _
                                 "SSE data #" & eventCount, payload

                    ElseIf eventCount = MAX_STREAM_TRACE_EVENTS + 1 Then

                        AddTrace ws, "STEP 4", "STREAM", _
                                 "Omit additional SSE trace entries", _
                                 "For readability, only the first " & MAX_STREAM_TRACE_EVENTS & _
                                 " SSE events are shown individually."

                    End If

                    If InStr(1, payload, """error""", vbTextCompare) > 0 Then
                        RaiseStreamingError payload
                    End If

                    If TryExtractJsonStringProperty(payload, "content", contentPart) Then

                        If Len(contentPart) > 0 Then
                            answerText = answerText & contentPart
                            ws.Range("B11").Value = answerText
                            DoEvents
                        End If

                    End If

                End If

            ElseIf readingHeaders Then

                responseHeaders = responseHeaders & lineText & vbCrLf

            ElseIf Len(lineText) > 0 Then

                nonSseOutput = nonSseOutput & lineText & vbCrLf

            End If

        Else

            DoEvents

        End If

        If exec.Status <> 0 And exec.StdOut.AtEndOfStream Then
            Exit Do
        End If

    Loop

    If Not exec.StdErr.AtEndOfStream Then
        stderrText = exec.StdErr.ReadAll
    End If

    If exec.ExitCode <> 0 Then
        Err.Raise vbObjectError + 2102, "SendOrcaRouterStreaming", _
                  "curl.exe exited with an error. ExitCode=" & exec.ExitCode & vbCrLf & stderrText
    End If

    If httpStatus <> 0 Then

        If httpStatus < 200 Or httpStatus >= 300 Then
            RaiseOrcaRouterHttpError httpStatus, responseHeaders, nonSseOutput
        End If

    End If

    AddTrace ws, "STEP 4", "RESPONSE", "Streaming receive completed", _
             "HTTP Status: " & IIf(httpStatus = 0, "(unknown)", CStr(httpStatus)) & vbCrLf & _
             "Events: " & eventCount & vbCrLf & _
             "Elapsed: " & Format$(ElapsedSeconds(startedAt), "0.000") & " sec" & vbCrLf & _
             "Headers:" & vbCrLf & responseHeaders & vbCrLf & _
             "Non-SSE body:" & vbCrLf & IIf(Len(nonSseOutput) = 0, "(empty)", nonSseOutput) & vbCrLf & _
             "curl stderr: " & IIf(Len(stderrText) = 0, "(empty)", stderrText)

    'STEP 5: Parse / process result.
    AddTrace ws, "STEP 5", "LOCAL", "Aggregate Streaming result", _
             "Answer length: " & Len(answerText)

    If Len(answerText) = 0 Then
        Err.Raise vbObjectError + 2103, "SendOrcaRouterStreaming", _
                  "Streaming completed but no answer text was captured. Check SSE data in the trace."
    End If

    'STEP 6: Update UI and trace.
    ws.Range("B11").Value = answerText

    AddTrace ws, "STEP 6", "LOCAL", "Display answer in worksheet", _
             "Mode: Streaming" & vbCrLf & _
             "Completed: True" & vbCrLf & _
             "Total elapsed: " & Format$(ElapsedSeconds(startedAt), "0.000") & " sec"

CleanExit:
    Set exec = Nothing
    Set shell = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next

    If Not ws Is Nothing Then

        ws.Range("B11").Value = "ERROR: " & errorDescription

        AddTrace ws, "ERROR", "ERROR", "Streaming error", _
                 "Err.Number: " & errorNumber & vbCrLf & _
                 "Err.Source: " & errorSource & vbCrLf & _
                 "Err.Description: " & errorDescription & vbCrLf & _
                 "curl stderr: " & IIf(Len(stderrText) = 0, "(not available)", stderrText)

    End If

    MsgBox "An error occurred during Streaming." & vbCrLf & _
           errorDescription & vbCrLf & vbCrLf & _
           "Check the worksheet trace for details.", _
           vbExclamation, _
           "OrcaRouter Streaming"

    On Error GoTo 0
    GoTo CleanExit

End Sub

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

    Dim startedAt As Double

    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME_ADV)

    apiKey = Trim$(CStr(ws.Range("B3").Value))
    model = Trim$(CStr(ws.Range("B4").Value))
    question = Trim$(CStr(ws.Range("B6").Value))

    ws.Range("B11").Value = vbNullString

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
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    SendJsonRequest httpRequest, apiKey, firstRequest, _
                    firstStatus, firstHeaders, firstResponse

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

    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    SendJsonRequest httpRequest, apiKey, secondRequest, _
                    secondStatus, secondHeaders, secondResponse

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
    ws.Range("B11").Value = assistantText

    AddTrace ws, "STEP 6", "LOCAL", "Display answer in worksheet", _
             "Mode: Tool Calling" & vbCrLf & _
             "Completed: True" & vbCrLf & _
             "Total elapsed: " & Format$(ElapsedSeconds(startedAt), "0.000") & " sec"

CleanExit:
    Set httpRequest = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next

    If Not ws Is Nothing Then

        ws.Range("B11").Value = "ERROR: " & errorDescription

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

    BuildStreamingRequestJson = _
        "{" & _
        """model"":""" & JsonEscape(model) & """," & _
        """messages"":[{" & _
            """role"":""user""," & _
            """content"":""" & JsonEscape(question) & """" & _
        "}]," & _
        """stream"":true," & _
        """stream_options"":{""include_usage"":true}" & _
        "}"

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
    result = result & "{"
    result = result & """role"":""user"","
    result = result & """content"":""" & JsonEscape(question) & """"
    result = result & "}"
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
    result = result & "{"
    result = result & """role"":""user"","
    result = result & """content"":""" & JsonEscape(question) & """"
    result = result & "},"
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

    AddTrace ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME_ADV), _
             "STEP 3", "REQUEST", "Send WinHTTP POST", _
             "Endpoint: " & API_ENDPOINT_ADV & vbCrLf & _
             "Authorization: Bearer " & MaskApiKey(apiKey)

    With httpRequest

        .Open "POST", API_ENDPOINT_ADV, False
        .SetTimeouts 10000, 10000, 30000, 60000
        .SetRequestHeader "Authorization", "Bearer " & apiKey
        .SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        .Send StringToUtf8Bytes(requestBody)

        httpStatus = .Status
        responseHeaders = .GetAllResponseHeaders
        responseText = .ResponseText

    End With

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
    guidance = BuildHttpGuidance(httpStatus, errorCode, retryAfter)

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
    ByVal retryAfter As String) As String

    Select Case httpStatus

        Case 401
            BuildHttpGuidance = _
                "The API key is invalid or the Authorization header is incorrect."

        Case 403

            Select Case errorCode

                Case "insufficient_user_quota"
                    BuildHttpGuidance = _
                        "Check workspace balance and member/agent budget."

                Case "pre_consume_token_quota_failed"
                    BuildHttpGuidance = _
                        "Check the quota limit assigned to this API key."

                Case "free_quota_exhausted"
                    BuildHttpGuidance = _
                        "No free model is currently available through the free router. Specify a paid model."

                Case Else
                    BuildHttpGuidance = _
                        "HTTP 403 has multiple possible causes. Check error.code and message."

            End Select

        Case 425
            BuildHttpGuidance = _
                "The selected model may not be available yet."

        Case 429

            If Len(retryAfter) > 0 Then
                BuildHttpGuidance = _
                    "Rate limit reached. Wait for Retry-After seconds before retrying."
            Else
                BuildHttpGuidance = _
                    "For a free-tier 429 without Retry-After, review prompt length and request shape."
            End If

        Case 503

            If errorCode = "model_not_found" Then
                BuildHttpGuidance = _
                    "Check whether the selected model is available for the current account."
            Else
                BuildHttpGuidance = _
                    "The service or upstream provider may be temporarily unavailable."
            End If

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

Private Function FindCurlPath() As String

    Dim systemRoot As String
    Dim candidate As String
    Dim shell As Object
    Dim exec As Object
    Dim outputText As String

    systemRoot = Environ$("SystemRoot")

    If Len(systemRoot) > 0 Then

        candidate = systemRoot & "\System32\curl.exe"

        If Len(Dir$(candidate)) > 0 Then
            FindCurlPath = candidate
            Exit Function
        End If

    End If

    On Error Resume Next

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("where curl.exe")
    outputText = Trim$(exec.StdOut.ReadLine)

    On Error GoTo 0

    If Len(outputText) > 0 Then
        FindCurlPath = outputText
    End If

    Set exec = Nothing
    Set shell = Nothing

End Function


Private Function ParseHttpStatusLine(ByVal statusLine As String) As Long

    Dim parts() As String

    parts = Split(Trim$(statusLine), " ")

    If UBound(parts) >= 1 Then

        If IsNumeric(parts(1)) Then
            ParseHttpStatusLine = CLng(parts(1))
        End If

    End If

End Function

Private Function BuildCurlConfig( _
    ByVal apiKey As String, _
    ByVal requestBody As String) As String

    BuildCurlConfig = _
        "url = """ & EscapeCurlConfigValue(API_ENDPOINT_ADV) & """" & vbCrLf & _
        "request = ""POST""" & vbCrLf & _
        "header = ""Authorization: Bearer " & _
            EscapeCurlConfigValue(apiKey) & """" & vbCrLf & _
        "header = ""Content-Type: application/json""" & vbCrLf & _
        "data = """ & EscapeCurlConfigValue(requestBody) & """" & vbCrLf

End Function

Private Function EscapeCurlConfigValue(ByVal value As String) As String

    Dim result As String

    result = Replace$(value, Chr$(92), Chr$(92) & Chr$(92))
    result = Replace$(result, Chr$(34), Chr$(92) & Chr$(34))
    result = Replace$(result, vbCr, Chr$(92) & "r")
    result = Replace$(result, vbLf, Chr$(92) & "n")

    EscapeCurlConfigValue = result

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
