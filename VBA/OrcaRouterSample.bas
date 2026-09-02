Attribute VB_Name = "OrcaRouterSample"
Option Explicit

'===============================================================================
' OrcaRouter API Learning Sample for Excel VBA
'
' Web / PowerShell / VBA use the same six conceptual steps.
'
' STEP 1 - Validate inputs
' STEP 2 - Build request
' STEP 3 - Send HTTP POST
' STEP 4 - Receive response
' STEP 5 - Parse assistant message
' STEP 6 - Update UI and trace
'
' This learning sample records HTTP status, raw response, errors,
' and more detail than a minimal production sample.
' The real API key is never written to the trace in plain text.
'===============================================================================

Private Const API_ENDPOINT As String = "https://api.orcarouter.ai/v1/chat/completions"
Private Const MODELS_ENDPOINT As String = "https://api.orcarouter.ai/v1/models"
Private Const API_KEY_PLACEHOLDER As String = "xxx-your-orcarouter-api-key-xxx"

'LOCAL TEST ONLY:
'LOCAL TEST ONLY: change only the next value if you temporarily embed a key.
'Never commit a real API key to a public GitHub repository.
Private Const DEFAULT_API_KEY As String = "xxx-your-orcarouter-api-key-xxx"

Private Const DEFAULT_MODEL As String = "orcarouter/free"
Private Const REFERRAL_URL As String = "https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d"
Private Const MAX_HISTORY_TURNS As Long = 10
Private Const SAMPLE_SHEET_NAME As String = "OrcaRouter Chat"
Private Const TRACE_HEADER_ROW As Long = 18
Private Const TRACE_FIRST_ROW As Long = 19
Private Const MAX_TRACE_TEXT As Long = 30000
Private Const RAW_JSON_MAX_TEXT As Long = 30000
Private Const CHAT_TOTAL_TIMEOUT_SECONDS As Long = 120
Private Const CHAT_WAIT_TRACE_INTERVAL_SECONDS As Long = 15
Private Const MODELS_TEST_TIMEOUT_MS As Long = 20000
Private Const UI_YIELD_INTERVAL_SECONDS As Double = 0.05
Private Const HTTP_POLL_INTERVAL_MS As Long = 50

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'DoEvents allows Excel to process pending UI messages while VBA is running.
'That also means the user can click the Send button again before the first
'request has finished. This flag prevents accidental re-entry / double sends.
Private requestInProgress As Boolean

'The Chat Completions API does not keep this history for us.
'The application stores the latest user/assistant pairs and sends them again.
Private conversationCount As Long
Private conversationUsers(1 To MAX_HISTORY_TURNS) As String
Private conversationAssistants(1 To MAX_HISTORY_TURNS) As String

Public Sub SetupOrcaRouterSample()

    Dim ws As Worksheet
    Dim sendButton As Shape
    Dim clearButton As Shape
    Dim applyPromptButton As Shape
    Dim newChatButton As Shape
    Dim listSeparator As String
    Dim workbookNameForOnAction As String
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler

    previousScreenUpdating = Application.ScreenUpdating

    Set ws = GetOrCreateSampleSheet()

    Application.ScreenUpdating = False

    ws.Cells.UnMerge
    ws.Cells.Clear

    listSeparator = Application.International(xlListSeparator)
    workbookNameForOnAction = Replace$(ThisWorkbook.Name, "'", "''")

    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop

    'Title
    With ws.Range("A1:H1")
        .Merge
        .Value = "OrcaRouter API Learning Sample - Excel VBA"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(23, 32, 51)
        .RowHeight = 32
    End With

    'First-run link. This is intentionally subtle and only explains where
    'a first-time user can create an OrcaRouter account/API key.
    ws.Range("A2").Value = "First run"
    With ws.Range("B2:H2")
        .Merge
        .Value = "Get an OrcaRouter API key (referral link)"
        .Font.Color = RGB(3, 105, 161)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    ws.Hyperlinks.Add Anchor:=ws.Range("B2"), _
                      Address:=REFERRAL_URL, _
                      TextToDisplay:="Get an OrcaRouter API key (referral link)"

    'API key
    ws.Range("A3").Value = "API Key"
    With ws.Range("B3:H3")
        .Merge
        .Value = DEFAULT_API_KEY
        .Interior.Color = RGB(251, 253, 255)
        .Borders.Color = RGB(217, 225, 236)
    End With

    'Model
    ws.Range("A4").Value = "Model"
    With ws.Range("B4:D4")
        .Merge
        .Value = DEFAULT_MODEL
        .Interior.Color = RGB(251, 253, 255)
        .Borders.Color = RGB(217, 225, 236)
    End With

    'Mode
    ws.Range("A5").Value = "Mode"
    With ws.Range("B5:D5")
        .Merge
        .Value = "Chat"
        .Interior.Color = RGB(251, 253, 255)
        .Borders.Color = RGB(217, 225, 236)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:="Chat" & listSeparator & _
                                 "Streaming" & listSeparator & _
                                 "Tool Calling"
    End With

    'Question
    ws.Range("A6").Value = "Question"
    With ws.Range("B6:H9")
        .Merge
        .Value = "Hello. Please describe OrcaRouter in one sentence."
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(251, 253, 255)
        .Borders.Color = RGB(217, 225, 236)
    End With

    'Prompt example selector
    ws.Range("A10").Value = "Prompt template (optional)"
    With ws.Range("B10:D10")
        .Merge
        .Value = "(select)"
        .Interior.Color = RGB(251, 253, 255)
        .Borders.Color = RGB(217, 225, 236)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:="(select)" & listSeparator & _
                                 "Summary" & listSeparator & _
                                 "Explain" & listSeparator & _
                                 "Code review" & listSeparator & _
                                 "JSON" & listSeparator & _
                                 "Translate"
    End With

    'Answer. Conversation history remains internal for the next API request.
    ws.Range("A11").Value = "Answer"
    With ws.Range("B11:H15")
        .Merge
        .Value = "The answer will appear here."
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(245, 247, 251)
        .Borders.Color = RGB(217, 225, 236)
    End With

    'Request status
    ws.Range("A16").Value = "Status"
    With ws.Range("B16:H16")
        .Merge
        .Value = "Ready"
        .Font.Color = RGB(71, 85, 105)
        .Interior.Color = RGB(248, 250, 252)
        .Borders.Color = RGB(217, 225, 236)
    End With

    'Raw JSON response area
    With ws.Range("J1:P1")
        .Merge
        .Value = "Raw JSON Response"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(23, 32, 51)
    End With

    With ws.Range("J2:P2")
        .Merge
        .Value = "No response yet."
        .Font.Bold = True
        .Font.Color = RGB(71, 85, 105)
    End With

    With ws.Range("J3:P15")
        .Merge
        .Value = vbNullString
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(250, 250, 250)
        .Borders.Color = RGB(217, 225, 236)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .NumberFormat = "@"
    End With

    'Developer information. Response JSON remains visible in J3:P15.
    With ws.Range("J17:P17")
        .Merge
        .Value = "Developer Information"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(23, 32, 51)
    End With

    ws.Range("J18").Value = "HTTP Status"
    ws.Range("L18").Value = "Elapsed"
    ws.Range("N18").Value = "Model"
    ws.Range("J19").Value = "Prompt Tokens"
    ws.Range("L19").Value = "Completion"
    ws.Range("N19").Value = "Total Tokens"
    ws.Range("J20").Value = "Cost"
    ws.Range("L20").Value = "History"
    ws.Range("J21").Value = "Request JSON"

    With ws.Range("J22:P30")
        .Merge
        .Value = "{}"
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(250, 250, 250)
        .Borders.Color = RGB(217, 225, 236)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .NumberFormat = "@"
    End With

    'Trace header
    ws.Range("A17").Value = "Processing steps / HTTP trace"
    ws.Range("A17").Font.Bold = True
    ws.Range("A17").Font.Size = 14

    ws.Cells(TRACE_HEADER_ROW, "A").Value = "Time"
    ws.Cells(TRACE_HEADER_ROW, "B").Value = "Step"
    ws.Cells(TRACE_HEADER_ROW, "C").Value = "Direction"
    ws.Cells(TRACE_HEADER_ROW, "D").Value = "Detail"
    ws.Cells(TRACE_HEADER_ROW, "E").Value = "Data"

    With ws.Range("A18:E18")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(15, 23, 42)
        .Borders.Color = RGB(217, 225, 236)
    End With

    'Column widths
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 90
    ws.Columns("F:H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 2
    ws.Columns("J:P").ColumnWidth = 12

    ws.Rows("2:5").RowHeight = 24
    ws.Rows("6:9").RowHeight = 26
    ws.Rows("10:10").RowHeight = 28
    ws.Rows("11:15").RowHeight = 26
    ws.Rows("16:16").RowHeight = 24

    ws.Range("A2:A17").Font.Bold = True
    ws.Range("A2:A17").Font.Color = RGB(71, 85, 105)

    'Send button
    Set sendButton = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("F4").Left, _
        ws.Range("F4").Top, _
        120, _
        30)

    With sendButton
        .Name = "btnSendOrcaRouter"
        .TextFrame2.TextRange.Text = "Send"
        .Fill.ForeColor.RGB = RGB(15, 23, 42)
        .Line.ForeColor.RGB = RGB(15, 23, 42)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .OnAction = "'" & workbookNameForOnAction & "'!SendOrcaRouterChat"
    End With

    'Apply prompt example button
    Set applyPromptButton = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("F10").Left, _
        ws.Range("F10").Top, _
        110, _
        26)

    With applyPromptButton
        .Name = "btnApplyPromptExample"
        .TextFrame2.TextRange.Text = "Insert prompt"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(217, 225, 236)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(23, 32, 51)
        .OnAction = "'" & workbookNameForOnAction & "'!ApplyOrcaRouterPromptExample"
    End With

    'New chat button
    Set newChatButton = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("G10").Left + 20, _
        ws.Range("G10").Top, _
        110, _
        26)

    With newChatButton
        .Name = "btnNewOrcaRouterChat"
        .TextFrame2.TextRange.Text = "New chat"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(217, 225, 236)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(23, 32, 51)
        .OnAction = "'" & workbookNameForOnAction & "'!NewOrcaRouterChat"
    End With

    'Clear trace button
    Set clearButton = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("G4").Left + 30, _
        ws.Range("G4").Top, _
        120, _
        30)

    With clearButton
        .Name = "btnClearOrcaRouterTrace"
        .TextFrame2.TextRange.Text = "Clear trace"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(217, 225, 236)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(23, 32, 51)
        .OnAction = "'" & workbookNameForOnAction & "'!ClearOrcaRouterTrace"
    End With

    ClearConversationHistoryState
    UpdateVbaHistoryStatus ws
    ResetOrcaRouterDeveloperInformation ws

    AddTrace ws, "READY", "LOCAL", "Created sample UI", _
             "Endpoint: " & API_ENDPOINT & vbCrLf & _
             "Model: " & DEFAULT_MODEL & vbCrLf & _
             "API Key: " & MaskApiKey(DEFAULT_API_KEY) & vbCrLf & _
             "Replace the dummy API key in cell B3 before running."

    ws.Activate
    ws.Range("B3").Select

CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Set newChatButton = Nothing
    Set applyPromptButton = Nothing
    Set clearButton = Nothing
    Set sendButton = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Failed to create the sample UI." & vbCrLf & _
           "Err.Number: " & Err.Number & vbCrLf & _
           "Err.Description: " & Err.Description, _
           vbExclamation, _
           "OrcaRouter Sample"
    Resume CleanExit

End Sub

Public Sub NewOrcaRouterChat()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    If requestInProgress Then
        MsgBox "A request is still running. Wait for it to finish before starting a new chat.", _
               vbInformation, _
               "OrcaRouter Sample"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)
    ClearConversationHistoryState
    ws.Range("B11").Value = "The answer will appear here."
    SetOrcaRouterUiStatus ws, "New chat - history cleared"
    UpdateVbaHistoryStatus ws

CleanExit:
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Failed to start a new chat." & vbCrLf & Err.Description, _
           vbExclamation, _
           "OrcaRouter Sample"
    Resume CleanExit

End Sub

Public Sub ApplyOrcaRouterPromptExample()

    Dim ws As Worksheet
    Dim exampleName As String
    Dim promptText As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)
    exampleName = Trim$(CStr(ws.Range("B10").Value))

    If Len(exampleName) = 0 Or exampleName = "(select)" Then
        MsgBox "Select a prompt template first.", _
               vbInformation, _
               "OrcaRouter Sample"
        GoTo CleanExit
    End If

    Select Case exampleName
        Case "Summary"
            promptText = "Summarize the following text in three concise bullet points:"
        Case "Explain"
            promptText = "Explain the following content for a beginner and define technical terms:"
        Case "Code review"
            promptText = "Review the following code. Explain issues, reasons, and an improved example:"
        Case "JSON"
            promptText = "Organize the following content and return JSON only:"
        Case "Translate"
            promptText = "Translate the following Japanese text into natural English:"
        Case Else
            MsgBox "Unknown prompt template: " & exampleName, _
                   vbExclamation, _
                   "OrcaRouter Sample"
            GoTo CleanExit
    End Select

    ws.Range("B6").Value = promptText & vbCrLf & vbCrLf & "Paste content here."

CleanExit:
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Failed to apply the prompt example." & vbCrLf & Err.Description, _
           vbExclamation, _
           "OrcaRouter Sample"
    Resume CleanExit

End Sub

Private Sub ClearConversationHistoryState()

    Dim i As Long

    conversationCount = 0

    For i = 1 To MAX_HISTORY_TURNS
        conversationUsers(i) = vbNullString
        conversationAssistants(i) = vbNullString
    Next i

End Sub

Public Sub CommitOrcaRouterConversationTurn(ByVal userText As String, ByVal assistantText As String)

    Dim i As Long

    If conversationCount >= MAX_HISTORY_TURNS Then
        For i = 1 To MAX_HISTORY_TURNS - 1
            conversationUsers(i) = conversationUsers(i + 1)
            conversationAssistants(i) = conversationAssistants(i + 1)
        Next i
        conversationCount = MAX_HISTORY_TURNS - 1
    End If

    conversationCount = conversationCount + 1
    conversationUsers(conversationCount) = userText
    conversationAssistants(conversationCount) = assistantText

End Sub

Public Function GetOrcaRouterAnswerDisplay( _
    Optional ByVal answerText As String = "") As String

    Dim result As String

    If Len(answerText) > 0 Then
        result = answerText
    ElseIf conversationCount > 0 Then
        result = conversationAssistants(conversationCount)
    Else
        result = "The answer will appear here."
    End If

    If Len(result) > RAW_JSON_MAX_TEXT Then
        result = Left$(result, RAW_JSON_MAX_TEXT - 40) & vbCrLf & _
                 "...(answer truncated for Excel cell limit)..."
    End If

    GetOrcaRouterAnswerDisplay = result

End Function

Public Sub SetOrcaRouterUiStatus( _
    ByVal ws As Worksheet, _
    ByVal statusMessage As String, _
    Optional ByVal isError As Boolean = False)

    ws.Range("B16").Value = statusMessage

    If isError Then
        ws.Range("B16").Font.Color = RGB(180, 35, 24)
    Else
        ws.Range("B16").Font.Color = RGB(71, 85, 105)
    End If

End Sub

Private Sub UpdateVbaHistoryStatus(ByVal ws As Worksheet)

    ws.Range("M20").Value = conversationCount & " / " & MAX_HISTORY_TURNS & " turns"

End Sub

Public Sub ResetOrcaRouterDeveloperInformation(ByVal ws As Worksheet)

    ws.Range("K18").Value = "-"
    ws.Range("M18").Value = "-"
    ws.Range("O18").Value = "-"
    ws.Range("K19").Value = "-"
    ws.Range("M19").Value = "-"
    ws.Range("O19").Value = "-"
    ws.Range("K20").Value = "(not returned)"
    ws.Range("J22").Value = "{}"
    UpdateVbaHistoryStatus ws

End Sub

Public Sub UpdateOrcaRouterDeveloperInformation( _
    ByVal ws As Worksheet, _
    ByVal httpStatus As Long, _
    ByVal elapsedSeconds As Double, _
    ByVal requestedModel As String, _
    ByVal requestJson As String, _
    ByVal responseJson As String, _
    Optional ByVal additionalResponseJson As String = "")

    Dim promptTokens As Double
    Dim completionTokens As Double
    Dim totalTokens As Double
    Dim costUsd As Double
    Dim actualModel As String

    If httpStatus > 0 Then
        ws.Range("K18").Value = httpStatus
    Else
        ws.Range("K18").Value = "(not available)"
    End If
    ws.Range("M18").Value = Format$(elapsedSeconds, "0.000") & " sec"

    actualModel = requestedModel
    If TryExtractAssistantStringProperty(responseJson, "model", actualModel) Then
        ws.Range("O18").Value = actualModel
    Else
        ws.Range("O18").Value = requestedModel
    End If

    If SumJsonNumberProperties( _
           responseJson, _
           additionalResponseJson, _
           "prompt_tokens", _
           promptTokens) Then
        ws.Range("K19").Value = promptTokens
    Else
        ws.Range("K19").Value = "-"
    End If

    If SumJsonNumberProperties( _
           responseJson, _
           additionalResponseJson, _
           "completion_tokens", _
           completionTokens) Then
        ws.Range("M19").Value = completionTokens
    Else
        ws.Range("M19").Value = "-"
    End If

    If SumJsonNumberProperties( _
           responseJson, _
           additionalResponseJson, _
           "total_tokens", _
           totalTokens) Then
        ws.Range("O19").Value = totalTokens
    Else
        ws.Range("O19").Value = "-"
    End If

    If SumJsonNumberProperties( _
           responseJson, _
           additionalResponseJson, _
           "cost_usd", _
           costUsd) Then
        ws.Range("K20").Value = "$" & Format$(costUsd, "0.000000")
    Else
        ws.Range("K20").Value = "(not returned)"
    End If

    If Len(requestJson) > 0 Then
        ws.Range("J22").Value = Left$(requestJson, RAW_JSON_MAX_TEXT)
    Else
        ws.Range("J22").Value = "{}"
    End If

    UpdateVbaHistoryStatus ws

End Sub

Private Function SumJsonNumberProperties( _
    ByVal primaryJson As String, _
    ByVal additionalJson As String, _
    ByVal propertyName As String, _
    ByRef totalValue As Double) As Boolean

    Dim oneValue As Double
    Dim foundValue As Boolean

    totalValue = 0
    foundValue = False

    If TryExtractJsonNumberProperty(primaryJson, propertyName, oneValue) Then
        totalValue = totalValue + oneValue
        foundValue = True
    End If

    oneValue = 0

    If Len(additionalJson) > 0 Then
        If TryExtractJsonNumberProperty(additionalJson, propertyName, oneValue) Then
            totalValue = totalValue + oneValue
            foundValue = True
        End If
    End If

    SumJsonNumberProperties = foundValue

End Function

Private Function TryExtractJsonNumberProperty( _
    ByVal jsonText As String, _
    ByVal propertyName As String, _
    ByRef numberValue As Double) As Boolean

    Dim regularExpression As Object
    Dim matches As Object
    Dim pattern As String

    Set regularExpression = CreateObject("VBScript.RegExp")

    pattern = Chr$(34) & propertyName & Chr$(34) & _
              Chr$(92) & "s*:" & Chr$(92) & "s*" & _
              "(-?[0-9]+(" & Chr$(92) & ".[0-9]+)?([Ee][+-]?[0-9]+)?)"

    With regularExpression
        .Global = False
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = pattern
    End With

    Set matches = regularExpression.Execute(jsonText)

    If matches.Count = 0 Then
        TryExtractJsonNumberProperty = False
    Else
        numberValue = Val(matches(0).SubMatches(0))
        TryExtractJsonNumberProperty = True
    End If

    Set matches = Nothing
    Set regularExpression = Nothing

End Function

Public Sub SendOrcaRouterChat()

    Dim ws As Worksheet
    Dim httpRequest As Object

    Dim apiKey As String
    Dim model As String
    Dim mode As String
    Dim question As String
    Dim requestBody As String

    Dim responseText As String
    Dim responseHeaders As String
    Dim assistantText As String

    Dim httpStatus As Long
    Dim startedAt As Double
    Dim elapsedTimeSeconds As Double

    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    If requestInProgress Then
        MsgBox "An OrcaRouter request is already running.", _
               vbInformation, _
               "OrcaRouter Sample"
        Exit Sub
    End If

    'From this point until CleanExit, block another Send action.
    'This is especially important because YieldToExcel calls DoEvents.
    requestInProgress = True

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)

    apiKey = Trim$(CStr(ws.Range("B3").Value))
    model = Trim$(CStr(ws.Range("B4").Value))
    mode = Trim$(CStr(ws.Range("B5").Value))
    question = Trim$(CStr(ws.Range("B6").Value))

    'Developer Information must describe this request attempt, not the previous one.
    ResetOrcaRouterDeveloperInformation ws

    If StrComp(mode, "Streaming", vbTextCompare) = 0 Then
        SendOrcaRouterStreaming
        GoTo CleanExit
    End If

    If StrComp(mode, "Tool Calling", vbTextCompare) = 0 Then
        SendOrcaRouterToolCalling
        GoTo CleanExit
    End If

    'Primary result is answer-only. Keep history internal and clear the
    'previous answer while waiting for the new response.
    ws.Range("B11").Value = vbNullString
    SetOrcaRouterUiStatus ws, "Waiting for OrcaRouter response..."
    PrepareRawResponse ws, "Raw JSON Response - " & mode

    'STEP 1: Validate inputs.
    AddTrace ws, "STEP 1", "LOCAL", "Validate inputs", _
             "API Key: " & MaskApiKey(apiKey) & vbCrLf & _
             "Model: " & model & vbCrLf & _
             "Mode: " & mode & vbCrLf & _
             "Question length: " & Len(question)

    If Len(apiKey) = 0 Or apiKey = API_KEY_PLACEHOLDER Or Left$(apiKey, 4) = "xxx-" Then
        Err.Raise vbObjectError + 1001, "SendOrcaRouterChat", _
                  "The API key is still the dummy value. Enter your OrcaRouter API key in cell B3."
    End If

    If Len(model) = 0 Then
        Err.Raise vbObjectError + 1002, "SendOrcaRouterChat", _
                  "Model is empty. Enter a model name in cell B4."
    End If

    If Len(question) = 0 Then
        Err.Raise vbObjectError + 1003, "SendOrcaRouterChat", _
                  "Question is empty. Enter a question in cell B6."
    End If

    'STEP 2: Build request.
    requestBody = BuildRequestJson(model, question)

    AddTrace ws, "STEP 2", "REQUEST", "Build HTTP request", _
             "Method: POST" & vbCrLf & _
             "Endpoint: " & API_ENDPOINT & vbCrLf & _
             "Authorization: Bearer " & MaskApiKey(apiKey) & vbCrLf & _
             "Content-Type: application/json; charset=utf-8" & vbCrLf & _
             "Body:" & vbCrLf & requestBody

    'STEP 3: Send HTTP POST.
    AddTrace ws, "STEP 3", "REQUEST", "Send POST to OrcaRouter", _
             "MSXML2.XMLHTTP.6.0 (asynchronous)" & vbCrLf & _
             "Polling interval: " & HTTP_POLL_INTERVAL_MS & " ms" & vbCrLf & _
             "Total wait limit: " & CHAT_TOTAL_TIMEOUT_SECONDS & " sec"

    startedAt = Timer

    'Use MSXML2.XMLHTTP.6.0 for normal desktop Excel requests.
    '
    'Why XMLHTTP here instead of WinHttp.WinHttpRequest?
    '  - PowerShell/.NET HttpClient was responding much faster on the same PC.
    '  - WinHTTP uses the machine-level WinHTTP networking/proxy stack.
    '  - XMLHTTP is intended for interactive desktop applications and follows
    '    the current-user networking/proxy settings more closely.
    '
    'This does not guarantee that every network will be faster, but it removes
    'the WinHTTP-specific path that was stalling before any HTTP status arrived.
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    With httpRequest

        'The third Open argument is the asynchronous flag.
        '
        'False = synchronous:
        '        Send blocks VBA until the HTTP request finishes or times out.
        '        While Send is blocking, VBA cannot reach DoEvents, so Excel can
        '        look frozen and Trace rows cannot repaint.
        '
        'True  = asynchronous:
        '        Send starts the request and returns control to VBA quickly.
        '        We then poll readyState, update Trace / StatusBar, and call
        '        YieldToExcel so Excel stays responsive while the request runs.
        .Open "POST", API_ENDPOINT, True
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "X-OrcaRouter-Include-Cost", "true"
        .Send StringToUtf8Bytes(requestBody)
    End With

    WaitForXmlHttpResponse _
        httpRequest, _
        ws, _
        "Waiting for OrcaRouter Chat response", _
        startedAt, _
        CHAT_TOTAL_TIMEOUT_SECONDS, _
        CHAT_WAIT_TRACE_INTERVAL_SECONDS

    httpStatus = httpRequest.Status
    responseHeaders = httpRequest.getAllResponseHeaders
    responseText = httpRequest.responseText

    DisplayRawResponse ws, responseText, "Raw JSON Response - Chat", httpStatus

    elapsedTimeSeconds = ElapsedSeconds(startedAt)

    'STEP 4: Receive response.
    AddTrace ws, "STEP 4", "RESPONSE", "Receive HTTP response", _
             "HTTP Status: " & httpStatus & vbCrLf & _
             "Elapsed: " & Format$(elapsedTimeSeconds, "0.000") & " sec" & vbCrLf & _
             "Headers:" & vbCrLf & responseHeaders & vbCrLf & _
             "Raw response:" & vbCrLf & responseText

    If httpStatus < 200 Or httpStatus >= 300 Then
        RaiseOrcaRouterHttpError httpStatus, responseHeaders, responseText
    End If

    'STEP 5: Parse assistant message.
    assistantText = ExtractAssistantContent(responseText)

    AddTrace ws, "STEP 5", "LOCAL", "Parse assistant message", _
             "Answer length: " & Len(assistantText)

    'STEP 6: Update UI and trace.
    CommitOrcaRouterConversationTurn question, assistantText
    ws.Range("B11").Value = GetOrcaRouterAnswerDisplay(assistantText)
    SetOrcaRouterUiStatus ws, "Completed"
    UpdateOrcaRouterDeveloperInformation _
        ws, _
        httpStatus, _
        elapsedTimeSeconds, _
        model, _
        requestBody, _
        responseText

    AddTrace ws, "STEP 6", "LOCAL", "Display answer in worksheet", _
             "Completed: True" & vbCrLf & _
             "History turns kept: " & conversationCount & " / " & MAX_HISTORY_TURNS & vbCrLf & _
             "Total elapsed: " & Format$(ElapsedSeconds(startedAt), "0.000") & " sec"

CleanExit:
    Application.StatusBar = False
    requestInProgress = False
    Set httpRequest = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next

    If Not ws Is Nothing Then

        'Do not commit failed turns. The Question remains in its own input area.
        ws.Range("B11").Value = _
            GetOrcaRouterAnswerDisplay("ERROR: " & errorDescription)
        SetOrcaRouterUiStatus ws, "Error - check Answer / Developer / Trace", True

        If startedAt > 0 Then
            elapsedTimeSeconds = ElapsedSeconds(startedAt)
        Else
            elapsedTimeSeconds = 0
        End If

        UpdateOrcaRouterDeveloperInformation _
            ws, _
            httpStatus, _
            elapsedTimeSeconds, _
            model, _
            requestBody, _
            responseText

        AddTrace ws, "ERROR", "ERROR", "An error occurred", _
                 "Err.Number: " & errorNumber & vbCrLf & _
                 "Err.Source: " & errorSource & vbCrLf & _
                 "Err.Description: " & errorDescription & vbCrLf & _
                 "HTTP Status: " & IIf(httpStatus = 0, "(no HTTP response)", CStr(httpStatus)) & vbCrLf & _
                 "Raw response:" & vbCrLf & IIf(Len(responseText) = 0, "(not available)", responseText)
    End If

    MsgBox "An error occurred while calling the OrcaRouter API." & vbCrLf & _
           "Err.Number: " & errorNumber & vbCrLf & _
           "Err.Description: " & errorDescription & vbCrLf & vbCrLf & _
           "Check the worksheet trace for details.", _
           vbExclamation, _
           "OrcaRouter Sample"

    On Error GoTo 0
    GoTo CleanExit

End Sub

Public Sub TestOrcaRouterConnection()

    Dim ws As Worksheet
    Dim httpRequest As Object

    Dim apiKey As String
    Dim responseText As String
    Dim responseHeaders As String
    Dim httpStatus As Long
    Dim startedAt As Double

    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)
    apiKey = Trim$(CStr(ws.Range("B3").Value))

    If Len(apiKey) = 0 Or apiKey = API_KEY_PLACEHOLDER Or Left$(apiKey, 4) = "xxx-" Then
        Err.Raise vbObjectError + 1201, "TestOrcaRouterConnection", _
                  "Enter the OrcaRouter API key in cell B3 first."
    End If

    AddTrace ws, "TEST", "REQUEST", "Test OrcaRouter API connectivity", _
             "GET " & MODELS_ENDPOINT & vbCrLf & _
             "Authorization: Bearer " & MaskApiKey(apiKey) & vbCrLf & _
             "Timeout: " & MODELS_TEST_TIMEOUT_MS / 1000 & " sec"

    startedAt = Timer
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")

    With httpRequest
        .Open "GET", MODELS_ENDPOINT, True
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send
    End With

    WaitForXmlHttpResponse _
        httpRequest, _
        ws, _
        "Waiting for OrcaRouter connectivity test", _
        startedAt, _
        MODELS_TEST_TIMEOUT_MS / 1000, _
        5

    httpStatus = httpRequest.Status
    responseHeaders = httpRequest.getAllResponseHeaders
    responseText = httpRequest.responseText

    AddTrace ws, "TEST", "RESPONSE", "Connectivity test response", _
             "HTTP Status: " & httpStatus & vbCrLf & _
             "Elapsed: " & Format$(ElapsedSeconds(startedAt), "0.000") & " sec" & vbCrLf & _
             "Headers:" & vbCrLf & responseHeaders & vbCrLf & _
             "Body preview:" & vbCrLf & Left$(responseText, 2000)

    If httpStatus >= 200 And httpStatus < 300 Then
        MsgBox "OrcaRouter connectivity and API-key authentication are working." & vbCrLf & _
               "The earlier Chat timeout is therefore more likely to be model/routing latency.", _
               vbInformation, _
               "OrcaRouter Connection Test"
    Else
        RaiseOrcaRouterHttpError httpStatus, responseHeaders, responseText
    End If

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
        AddTrace ws, "TEST", "ERROR", "Connectivity test failed", _
                 "Err.Number: " & errorNumber & vbCrLf & _
                 "Err.Source: " & errorSource & vbCrLf & _
                 "Err.Description: " & errorDescription
    End If

    MsgBox "OrcaRouter connection test failed." & vbCrLf & _
           "Err.Number: " & errorNumber & vbCrLf & _
           "Err.Description: " & errorDescription & vbCrLf & vbCrLf & _
           "Check the worksheet trace for details.", _
           vbExclamation, _
           "OrcaRouter Connection Test"

    On Error GoTo 0
    GoTo CleanExit

End Sub

Public Sub RunOrcaRouterVbaSelfTests()

    Dim originalText As String
    Dim encodedText As String
    Dim decodedText As String
    Dim requestJson As String
    Dim displayText As String
    Dim utf8Bytes As Variant
    Dim byteCount As Long

    On Error GoTo ErrorHandler

    originalText = "Line1" & vbLf & "Quote:" & Chr$(34) & " Backslash:" & Chr$(92)
    encodedText = JsonEscape(originalText)
    decodedText = JsonUnescape(encodedText)

    If decodedText <> originalText Then
        Err.Raise vbObjectError + 3001, "RunOrcaRouterVbaSelfTests", _
                  "JsonEscape / JsonUnescape round-trip failed."
    End If

    ClearConversationHistoryState
    requestJson = BuildRequestJson("orcarouter/free", "hello")

    If requestJson <> _
       "{""model"":""orcarouter/free"",""messages"":[{""role"":""user"",""content"":""hello""}]}" Then

        Err.Raise vbObjectError + 3002, "RunOrcaRouterVbaSelfTests", _
                  "BuildRequestJson produced an unexpected result."
    End If

    CommitOrcaRouterConversationTurn "first question", "first answer"
    requestJson = BuildRequestJson("orcarouter/free", "second question")

    If InStr(1, requestJson, """role"":""assistant""", vbBinaryCompare) = 0 Or _
       InStr(1, requestJson, "first answer", vbBinaryCompare) = 0 Or _
       InStr(1, requestJson, "second question", vbBinaryCompare) = 0 Then

        Err.Raise vbObjectError + 3005, "RunOrcaRouterVbaSelfTests", _
                  "Conversation history was not included in BuildRequestJson."
    End If

    'Advanced builders must reuse the same committed history.
    RunOrcaRouterAdvancedSelfTests True

    displayText = GetOrcaRouterAnswerDisplay("ERROR: synthetic failure")

    If InStr(1, displayText, "failed question", vbBinaryCompare) > 0 Or _
       InStr(1, displayText, "ERROR: synthetic failure", vbBinaryCompare) = 0 Then

        Err.Raise vbObjectError + 3006, "RunOrcaRouterVbaSelfTests", _
                  "Answer display must show error only and must not repeat Question."
    End If

    displayText = GetOrcaRouterAnswerDisplay()

    If displayText <> "first answer" Then
        Err.Raise vbObjectError + 3007, "RunOrcaRouterVbaSelfTests", _
                  "Answer display must show only the latest committed assistant response."
    End If

    ClearConversationHistoryState

    If MaskApiKey("1234567890") <> "1234...7890" Then
        Err.Raise vbObjectError + 3003, "RunOrcaRouterVbaSelfTests", _
                  "MaskApiKey produced an unexpected result."
    End If

    utf8Bytes = StringToUtf8Bytes("ABC")
    byteCount = UBound(utf8Bytes) - LBound(utf8Bytes) + 1

    If byteCount <> 3 Then
        Err.Raise vbObjectError + 3004, "RunOrcaRouterVbaSelfTests", _
                  "UTF-8 conversion produced an unexpected byte count."
    End If

    MsgBox "All local VBA self-tests passed.", _
           vbInformation, _
           "OrcaRouter VBA Self-Test"

    Exit Sub

ErrorHandler:
    MsgBox "VBA self-test failed." & vbCrLf & _
           "Err.Number: " & Err.Number & vbCrLf & _
           "Err.Source: " & Err.Source & vbCrLf & _
           "Err.Description: " & Err.Description, _
           vbExclamation, _
           "OrcaRouter VBA Self-Test"

End Sub

Public Sub ClearOrcaRouterTrace()

    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)

    ws.Range("A" & TRACE_FIRST_ROW & ":E" & ws.Rows.Count).ClearContents

    AddTrace ws, "READY", "LOCAL", "Trace cleared", _
             "A new trace will be recorded on the next request."

CleanExit:
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Failed to clear the trace." & vbCrLf & Err.Description, _
           vbExclamation, _
           "OrcaRouter Sample"
    Resume CleanExit

End Sub

Private Function GetOrCreateSampleSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SAMPLE_SHEET_NAME
    End If

    Set GetOrCreateSampleSheet = ws

End Function

Public Function BuildOrcaRouterConversationMessageItemsJson( _
    ByVal question As String) As String

    Dim i As Long
    Dim result As String

    For i = 1 To conversationCount
        If Len(result) > 0 Then
            result = result & ","
        End If

        result = result & "{""role"":""user"",""content"":""" & _
                 JsonEscape(conversationUsers(i)) & """},"
        result = result & "{""role"":""assistant"",""content"":""" & _
                 JsonEscape(conversationAssistants(i)) & """}"
    Next i

    If Len(result) > 0 Then
        result = result & ","
    End If

    result = result & "{""role"":""user"",""content"":""" & _
             JsonEscape(question) & """}"

    BuildOrcaRouterConversationMessageItemsJson = result

End Function

Private Function BuildRequestJson(ByVal model As String, ByVal question As String) As String

    Dim result As String

    result = "{"
    result = result & """model"":""" & JsonEscape(model) & ""","
    result = result & """messages"":["
    result = result & BuildOrcaRouterConversationMessageItemsJson(question)
    result = result & "]"
    result = result & "}"

    BuildRequestJson = result

End Function

Public Function JsonEscape(ByVal value As String) As String

    Dim i As Long
    Dim oneCharacter As String
    Dim codeUnit As Long
    Dim result As String

    For i = 1 To Len(value)

        oneCharacter = Mid$(value, i, 1)
        codeUnit = AscW(oneCharacter)

        Select Case codeUnit

            Case 34
                result = result & Chr$(92) & Chr$(34)

            Case 92
                result = result & Chr$(92) & Chr$(92)

            Case 8
                result = result & Chr$(92) & "b"

            Case 9
                result = result & Chr$(92) & "t"

            Case 10
                result = result & Chr$(92) & "n"

            Case 12
                result = result & Chr$(92) & "f"

            Case 13
                result = result & Chr$(92) & "r"

            Case 0 To 31
                result = result & Chr$(92) & "u" & Right$("0000" & Hex$(codeUnit), 4)

            Case Else
                result = result & oneCharacter

        End Select

    Next i

    JsonEscape = result

End Function

Public Function ExtractAssistantContent(ByVal responseJson As String) As String

    Dim choicesPosition As Long
    Dim messagePosition As Long
    Dim messageJson As String
    Dim decodedValue As String

    'Expected OpenAI-compatible shape:
    'choices[0].message.content
    '
    'Scope the search to the "message" section instead of taking the first
    'property named "content" from the whole response. This reduces the chance
    'of accidentally reading another content field if the response grows.
    choicesPosition = InStr(1, responseJson, """choices""", vbTextCompare)

    If choicesPosition = 0 Then
        Err.Raise vbObjectError + 1101, "ExtractAssistantContent", _
                  "Could not find choices in the JSON response."
    End If

    messagePosition = InStr(choicesPosition, responseJson, """message""", vbTextCompare)

    If messagePosition = 0 Then
        Err.Raise vbObjectError + 1102, "ExtractAssistantContent", _
                  "Could not find choices[0].message in the JSON response."
    End If

    messageJson = Mid$(responseJson, messagePosition)

    'Normal Chat Completions usually returns content as one JSON string.
    If TryExtractAssistantStringProperty(messageJson, "content", decodedValue) Then
        ExtractAssistantContent = decodedValue
        Exit Function
    End If

    'Some OpenAI-compatible responses can represent message content as an
    'array of parts, for example: [{"type":"text","text":"hello"}].
    'For this lightweight learning sample, use the first text string as a
    'fallback so Answer is still populated for that common shape.
    decodedValue = ExtractAssistantTextParts(messageJson)

    If Len(decodedValue) > 0 Then
        ExtractAssistantContent = decodedValue
        Exit Function
    End If

    Err.Raise vbObjectError + 1103, "ExtractAssistantContent", _
              "Could not extract choices[0].message.content. See Raw JSON Response."

End Function

Private Function TryExtractAssistantStringProperty( _
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
        TryExtractAssistantStringProperty = False
    Else
        decodedValue = JsonUnescape(matches(0).SubMatches(0))
        TryExtractAssistantStringProperty = True
    End If

    Set matches = Nothing
    Set regularExpression = Nothing

End Function

Private Function ExtractAssistantTextParts(ByVal messageJson As String) As String

    Dim regularExpression As Object
    Dim matches As Object
    Dim oneMatch As Object
    Dim pattern As String
    Dim result As String
    Dim decodedPart As String

    Set regularExpression = CreateObject("VBScript.RegExp")

    pattern = _
        Chr$(34) & "text" & Chr$(34) & _
        Chr$(92) & "s*:" & Chr$(92) & "s*" & _
        Chr$(34) & _
        "((" & Chr$(92) & Chr$(92) & ".|[^" & _
        Chr$(34) & Chr$(92) & Chr$(92) & "])*)" & _
        Chr$(34)

    With regularExpression
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = pattern
    End With

    Set matches = regularExpression.Execute(messageJson)

    For Each oneMatch In matches

        decodedPart = JsonUnescape(oneMatch.SubMatches(0))

        If Len(decodedPart) > 0 Then

            If Len(result) > 0 Then
                result = result & vbCrLf
            End If

            result = result & decodedPart

        End If

    Next oneMatch

    ExtractAssistantTextParts = result

    Set matches = Nothing
    Set regularExpression = Nothing

End Function

Public Function JsonUnescape(ByVal encodedText As String) As String

    Dim i As Long
    Dim oneCharacter As String
    Dim escapeCharacter As String
    Dim hexText As String
    Dim codeUnit As Long
    Dim result As String

    i = 1

    Do While i <= Len(encodedText)

        oneCharacter = Mid$(encodedText, i, 1)

        If oneCharacter <> Chr$(92) Then

            result = result & oneCharacter
            i = i + 1

        Else

            If i = Len(encodedText) Then
                result = result & Chr$(92)
                Exit Do
            End If

            escapeCharacter = Mid$(encodedText, i + 1, 1)

            Select Case escapeCharacter

                Case Chr$(34)
                    result = result & Chr$(34)
                    i = i + 2

                Case Chr$(92)
                    result = result & Chr$(92)
                    i = i + 2

                Case "/"
                    result = result & "/"
                    i = i + 2

                Case "b"
                    result = result & Chr$(8)
                    i = i + 2

                Case "f"
                    result = result & Chr$(12)
                    i = i + 2

                Case "n"
                    result = result & vbLf
                    i = i + 2

                Case "r"
                    result = result & vbCr
                    i = i + 2

                Case "t"
                    result = result & vbTab
                    i = i + 2

                Case "u"

                    If i + 5 > Len(encodedText) Then
                        Err.Raise vbObjectError + 1102, "JsonUnescape", _
                                  "Invalid Unicode escape sequence detected."
                    End If

                    hexText = Mid$(encodedText, i + 2, 4)
                    codeUnit = CLng("&H" & hexText)

                    If codeUnit > &H7FFF Then
                        codeUnit = codeUnit - &H10000
                    End If

                    result = result & ChrW$(codeUnit)
                    i = i + 6

                Case Else
                    result = result & escapeCharacter
                    i = i + 2

            End Select

        End If

    Loop

    JsonUnescape = result

End Function


Public Function StringToUtf8Bytes(ByVal value As String) As Variant

    Dim stream As Object

    'Use ADODB.Stream via late binding to convert a VBA String to UTF-8 bytes.
    'This keeps non-ASCII text safe when sending JSON.
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText value

        .Position = 0
        .Type = 1

        'Skip the 3-byte UTF-8 BOM because it is unnecessary in the HTTP body.
        .Position = 3

        StringToUtf8Bytes = .Read
        .Close
    End With

    Set stream = Nothing

End Function

Public Function MaskApiKey(ByVal apiKey As String) As String

    If Len(apiKey) = 0 Then
        MaskApiKey = "(empty)"
        Exit Function
    End If

    If Len(apiKey) <= 8 Then
        MaskApiKey = "********"
        Exit Function
    End If

    MaskApiKey = Left$(apiKey, 4) & "..." & Right$(apiKey, 4)

End Function

Public Function ElapsedSeconds(ByVal startedAt As Double) As Double

    Dim currentTime As Double

    currentTime = Timer

    If currentTime >= startedAt Then
        ElapsedSeconds = currentTime - startedAt
    Else
        'Timer resets at midnight.
        ElapsedSeconds = (86400# - startedAt) + currentTime
    End If

End Function

Public Sub PrepareRawResponse( _
    ByVal ws As Worksheet, _
    ByVal caption As String)

    ws.Range("J1").Value = caption
    ws.Range("J2").Value = "Waiting for HTTP response..."
    ws.Range("J3").Value = vbNullString

    YieldToExcel True

End Sub

Public Sub DisplayRawResponse( _
    ByVal ws As Worksheet, _
    ByVal responseText As String, _
    ByVal caption As String, _
    Optional ByVal httpStatus As Long = 0)

    Dim statusText As String

    ws.Range("J1").Value = caption

    If httpStatus = 0 Then
        statusText = "HTTP Status: (not available)"
    Else
        statusText = "HTTP Status: " & CStr(httpStatus)
    End If

    statusText = statusText & " / Chars: " & Len(responseText)

    ws.Range("J2").Value = statusText
    ws.Range("J3").Value = LimitRawJsonText(responseText)

    YieldToExcel True

End Sub

Private Function LimitRawJsonText(ByVal value As String) As String

    If Len(value) <= RAW_JSON_MAX_TEXT Then
        LimitRawJsonText = value
    Else
        LimitRawJsonText = Left$(value, RAW_JSON_MAX_TEXT) & vbCrLf & _
                           "...(truncated for Excel cell limit / readability)"
    End If

End Function

Public Sub AddTrace( _
    ByVal ws As Worksheet, _
    ByVal stepName As String, _
    ByVal direction As String, _
    ByVal detail As String, _
    Optional ByVal data As String = vbNullString)

    Dim nextRow As Long

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    If nextRow < TRACE_FIRST_ROW Then
        nextRow = TRACE_FIRST_ROW
    End If

    ws.Cells(nextRow, "A").Value = Format$(Now, "hh:mm:ss")
    ws.Cells(nextRow, "B").Value = stepName
    ws.Cells(nextRow, "C").Value = direction
    ws.Cells(nextRow, "D").Value = detail
    ws.Cells(nextRow, "E").Value = LimitTraceText(data)

    With ws.Range("A" & nextRow & ":E" & nextRow)
        .VerticalAlignment = xlTop
        .WrapText = True
        .Borders.Color = RGB(217, 225, 236)
    End With

    ws.Cells(nextRow, "B").Font.Bold = True

    Select Case direction

        Case "REQUEST"
            ws.Cells(nextRow, "C").Font.Color = RGB(3, 105, 161)

        Case "RESPONSE"
            ws.Cells(nextRow, "C").Font.Color = RGB(15, 118, 110)

        Case "ERROR"
            ws.Cells(nextRow, "C").Font.Color = RGB(180, 35, 24)

        Case Else
            ws.Cells(nextRow, "C").Font.Color = RGB(71, 85, 105)

    End Select

    If Len(data) > 0 Then
        ws.Rows(nextRow).RowHeight = 72
    End If

    YieldToExcel

End Sub

Public Sub YieldToExcel(Optional ByVal force As Boolean = False)

    'YieldToExcel is a small wrapper around VBA's standard DoEvents statement.
    '
    'Why not call DoEvents everywhere?
    '  - Calling it too often adds overhead.
    '  - DoEvents makes the application re-entrant: button clicks and other
    '    queued Excel events can run while this macro is still in progress.
    '  - A central helper makes the intended UI-yield points easy to audit.
    '
    'The helper therefore throttles DoEvents to about once every 0.05 seconds.
    'That is frequent enough for screen repainting and window movement without
    'spending every loop iteration processing the Windows message queue.

    Static lastYieldAt As Double

    Dim currentTime As Double
    Dim elapsed As Double

    currentTime = Timer

    If currentTime >= lastYieldAt Then
        elapsed = currentTime - lastYieldAt
    Else
        elapsed = (86400# - lastYieldAt) + currentTime
    End If

    If force Or elapsed >= UI_YIELD_INTERVAL_SECONDS Then
        DoEvents
        lastYieldAt = currentTime
    End If

End Sub

Public Sub WaitForXmlHttpResponse( _
    ByVal httpRequest As Object, _
    ByVal ws As Worksheet, _
    ByVal waitMessage As String, _
    ByVal startedAt As Double, _
    ByVal totalTimeoutSeconds As Double, _
    Optional ByVal traceIntervalSeconds As Double = 15)

    Dim elapsedSecondsValue As Double
    Dim nextTraceAt As Double

    'MSXML2.XMLHTTP uses readyState to report asynchronous progress.
    '
    'readyState = 4 means the HTTP operation is complete.
    'Unlike WinHttp.WinHttpRequest, XMLHTTP does not expose WaitForResponse,
    'so this helper polls readyState at a small interval.
    '
    'The short Sleep prevents a tight CPU-burning loop. After each short pause
    'YieldToExcel calls DoEvents, which lets Excel repaint the Trace, move its
    'window, and remain responsive while the request is still running.
    nextTraceAt = traceIntervalSeconds

    Do While httpRequest.readyState <> 4

        elapsedSecondsValue = ElapsedSeconds(startedAt)

        Application.StatusBar = _
            waitMessage & " - " & Format$(elapsedSecondsValue, "0") & " sec"

        If elapsedSecondsValue >= nextTraceAt Then
            AddTrace ws, "WAIT", "WAIT", waitMessage, _
                     "Elapsed: " & Format$(elapsedSecondsValue, "0") & " sec" & vbCrLf & _
                     "readyState: " & CStr(httpRequest.readyState)

            nextTraceAt = nextTraceAt + traceIntervalSeconds
        End If

        If elapsedSecondsValue >= totalTimeoutSeconds Then
            httpRequest.abort
            Application.StatusBar = False

            Err.Raise vbObjectError + 1301, "WaitForXmlHttpResponse", _
                      waitMessage & " timed out after " & _
                      Format$(totalTimeoutSeconds, "0") & " seconds."
        End If

        Sleep HTTP_POLL_INTERVAL_MS
        YieldToExcel True

    Loop

    Application.StatusBar = False

End Sub

Private Function LimitTraceText(ByVal value As String) As String

    If Len(value) <= MAX_TRACE_TEXT Then
        LimitTraceText = value
    Else
        LimitTraceText = Left$(value, MAX_TRACE_TEXT) & vbCrLf & _
                         "...(truncated for Excel cell limit / readability)"
    End If

End Function
