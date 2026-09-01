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
Private Const API_KEY_PLACEHOLDER As String = "xxx-your-orcarouter-api-key-xxx"

'LOCAL TEST ONLY:
'LOCAL TEST ONLY: change only the next value if you temporarily embed a key.
'Never commit a real API key to a public GitHub repository.
Private Const DEFAULT_API_KEY As String = "xxx-your-orcarouter-api-key-xxx"

Private Const DEFAULT_MODEL As String = "orcarouter/free"
Private Const SAMPLE_SHEET_NAME As String = "OrcaRouter Chat"
Private Const TRACE_HEADER_ROW As Long = 18
Private Const TRACE_FIRST_ROW As Long = 19
Private Const MAX_TRACE_TEXT As Long = 30000

Public Sub SetupOrcaRouterSample()

    Dim ws As Worksheet
    Dim sendButton As Shape
    Dim clearButton As Shape
    Dim listSeparator As String
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler

    Set ws = GetOrCreateSampleSheet()

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ws.Cells.UnMerge
    ws.Cells.Clear

    listSeparator = Application.International(xlListSeparator)

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

    'Answer
    ws.Range("A11").Value = "Answer"
    With ws.Range("B11:H15")
        .Merge
        .Value = "The answer will appear here."
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(245, 247, 251)
        .Borders.Color = RGB(217, 225, 236)
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

    ws.Rows("3:5").RowHeight = 24
    ws.Rows("6:9").RowHeight = 26
    ws.Rows("11:15").RowHeight = 26

    ws.Range("A3:A17").Font.Bold = True
    ws.Range("A3:A17").Font.Color = RGB(71, 85, 105)

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
        .OnAction = "'" & ThisWorkbook.Name & "'!SendOrcaRouterChat"
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
        .OnAction = "'" & ThisWorkbook.Name & "'!ClearOrcaRouterTrace"
    End With

    AddTrace ws, "READY", "LOCAL", "Created sample UI", _
             "Endpoint: " & API_ENDPOINT & vbCrLf & _
             "Model: " & DEFAULT_MODEL & vbCrLf & _
             "API Key: " & MaskApiKey(DEFAULT_API_KEY) & vbCrLf & _
             "Replace the dummy API key in cell B3 before running."

    ws.Activate
    ws.Range("B3").Select

CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
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

    Set ws = ThisWorkbook.Worksheets(SAMPLE_SHEET_NAME)

    apiKey = Trim$(CStr(ws.Range("B3").Value))
    model = Trim$(CStr(ws.Range("B4").Value))
    mode = Trim$(CStr(ws.Range("B5").Value))
    question = Trim$(CStr(ws.Range("B6").Value))

    If StrComp(mode, "Streaming", vbTextCompare) = 0 Then
        SendOrcaRouterStreaming
        GoTo CleanExit
    End If

    If StrComp(mode, "Tool Calling", vbTextCompare) = 0 Then
        SendOrcaRouterToolCalling
        GoTo CleanExit
    End If

    ws.Range("B11").Value = vbNullString

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
             "WinHttp.WinHttpRequest.5.1" & vbCrLf & _
             "Timeouts(ms): Resolve=10000, Connect=10000, Send=30000, Receive=60000"

    startedAt = Timer

    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    With httpRequest
        .Open "POST", API_ENDPOINT, False
        .SetTimeouts 10000, 10000, 30000, 60000
        .SetRequestHeader "Authorization", "Bearer " & apiKey
        .SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        .Send StringToUtf8Bytes(requestBody)

        httpStatus = .Status
        responseHeaders = .GetAllResponseHeaders
        responseText = .ResponseText
    End With

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
    ws.Range("B11").Value = assistantText

    AddTrace ws, "STEP 6", "LOCAL", "Display answer in worksheet", _
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

Private Function BuildRequestJson(ByVal model As String, ByVal question As String) As String

    BuildRequestJson = _
        "{" & _
        """model"":""" & JsonEscape(model) & """," & _
        """messages"":[{" & _
            """role"":""user""," & _
            """content"":""" & JsonEscape(question) & """" & _
        "}]" & _
        "}"

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

    Dim regularExpression As Object
    Dim matches As Object
    Dim encodedContent As String
    Dim pattern As String

    Set regularExpression = CreateObject("VBScript.RegExp")

    'Expected OpenAI-compatible shape:
    'choices[0].message.content
    '
    'This lightweight sample extracts the first JSON string property named "content".
    'For production code, use a full JSON parser.
    pattern = _
        Chr$(34) & "content" & Chr$(34) & _
        Chr$(92) & "s*:" & Chr$(92) & "s*" & _
        Chr$(34) & _
        "((" & Chr$(92) & Chr$(92) & ".|[^" & Chr$(34) & Chr$(92) & Chr$(92) & "])*)" & _
        Chr$(34)

    With regularExpression
        .Global = False
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = pattern
    End With

    Set matches = regularExpression.Execute(responseJson)

    If matches.Count = 0 Then
        Err.Raise vbObjectError + 1101, "ExtractAssistantContent", _
                  "Could not extract choices[0].message.content. Check the raw response in the trace."
    End If

    encodedContent = matches(0).SubMatches(0)
    ExtractAssistantContent = JsonUnescape(encodedContent)

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

End Sub

Private Function LimitTraceText(ByVal value As String) As String

    If Len(value) <= MAX_TRACE_TEXT Then
        LimitTraceText = value
    Else
        LimitTraceText = Left$(value, MAX_TRACE_TEXT) & vbCrLf & _
                         "...(truncated for Excel cell limit / readability)"
    End If

End Function
