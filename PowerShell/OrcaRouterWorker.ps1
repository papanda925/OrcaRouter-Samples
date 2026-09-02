#requires -Version 5.1
<#
.SYNOPSIS
    Background worker for the OrcaRouter PowerShell/WPF sample.

.DESCRIPTION
    Runs OrcaRouter HTTP work in a separate PowerShell runspace.
    It never touches WPF controls directly. Progress is sent to the UI through
    a ConcurrentQueue as Trace / Answer / Completed / Error events.
#>

param(
    [switch]$SelfTest,
    [switch]$PipelineSelfTest,
    [string]$ApiKey = '',
    [string]$Model = 'orcarouter/free',
    [string]$Question = '',
    [ValidateSet('Chat', 'Streaming', 'Tool Calling')]
    [string]$Mode = 'Chat',
    [object[]]$History = @(),
    [Parameter(Mandatory = $true)]
    [System.Collections.Concurrent.ConcurrentQueue[object]]$EventQueue
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Net.Http

$apiEndpoint = 'https://api.orcarouter.ai/v1/chat/completions'
$requestTimeoutSeconds = 120
$maxStreamTraceEvents = 50

# Keep the latest request/response metadata so an Error event can still
# populate the Answer and Developer tabs instead of losing diagnostic data.
$script:lastRequest = $null
$script:lastResponse = $null
$script:lastHttpStatus = $null
$script:lastUsage = $null
$script:lastActualModel = $Model

# Normalize History once. PowerShell functions can emit $null/scalar/array values
# differently across Windows PowerShell 5.1 call paths. Never rely on a dynamic
# .Count property for request history.
$historyItems = @()
$historyTurnCount = 0

foreach ($turn in $historyItems) {
    if ($null -eq $turn) { continue }
    $historyItems += $turn
    $historyTurnCount += 1
}

function Add-WorkerEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [hashtable]$Values
    )

    $eventData = [ordered]@{ Type = $Type }

    if ($null -ne $Values) {
        foreach ($key in $Values.Keys) {
            $eventData[$key] = $Values[$key]
        }
    }

    $EventQueue.Enqueue([pscustomobject]$eventData)
}

function Add-WorkerTrace {
    param(
        [string]$Step,
        [string]$Direction,
        [string]$Title,
        $Data
    )

    Add-WorkerEvent -Type 'Trace' -Values @{
        Step = $Step
        Direction = $Direction
        Title = $Title
        Data = $Data
    }
}


if ($SelfTest) {
    Add-WorkerEvent -Type 'Trace' -Values @{
        Step = 'SELFTEST'
        Direction = 'LOCAL'
        Title = 'Background runspace self-test started'
        Data = @{ ThreadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId }
    }

    Start-Sleep -Milliseconds 150

    Add-WorkerEvent -Type 'Answer' -Values @{
        Text = 'Background runspace self-test answer'
    }

    Add-WorkerEvent -Type 'Completed' -Values @{
        Answer = 'Background runspace self-test answer'
        TotalElapsedMs = 150
    }

    return
}

function Mask-ApiKey {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '(empty)' }
    if ($Value.Length -le 8) { return '********' }

    return '{0}...{1}' -f $Value.Substring(0, 4), $Value.Substring($Value.Length - 4)
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $Object) { return $null }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }

    return $property.Value
}

function New-ConversationMessages {
    param([string]$CurrentQuestion)

    $messages = @()

    foreach ($turn in $historyItems) {
        if ($null -eq $turn) { continue }

        $userText = Get-PropertyValue -Object $turn -Name 'User'
        $assistantText = Get-PropertyValue -Object $turn -Name 'Assistant'

        if (-not [string]::IsNullOrWhiteSpace([string]$userText)) {
            $messages += [ordered]@{
                role = 'user'
                content = [string]$userText
            }
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$assistantText)) {
            $messages += [ordered]@{
                role = 'assistant'
                content = [string]$assistantText
            }
        }
    }

    $messages += [ordered]@{
        role = 'user'
        content = $CurrentQuestion
    }

    return $messages
}

function Merge-Usage {
    param(
        $FirstUsage,
        $SecondUsage
    )

    if ($null -eq $FirstUsage -and $null -eq $SecondUsage) {
        return $null
    }

    $result = [ordered]@{
        prompt_tokens = 0
        completion_tokens = 0
        total_tokens = 0
    }

    $cost = 0.0
    $hasCost = $false

    foreach ($usage in @($FirstUsage, $SecondUsage)) {
        if ($null -eq $usage) { continue }

        foreach ($name in @('prompt_tokens', 'completion_tokens', 'total_tokens')) {
            $value = Get-PropertyValue -Object $usage -Name $name
            if ($null -ne $value) {
                $result[$name] += [long]$value
            }
        }

        $costValue = Get-PropertyValue -Object $usage -Name 'cost_usd'
        if ($null -ne $costValue) {
            $cost += [double]$costValue
            $hasCost = $true
        }
    }

    if ($hasCost) {
        $result['cost_usd'] = $cost
    }

    return [pscustomobject]$result
}

function Get-ResponseHeaders {
    param([System.Net.Http.HttpResponseMessage]$Response)

    $headers = [ordered]@{}

    foreach ($header in $Response.Headers) {
        $headers[$header.Key] = ($header.Value -join ', ')
    }

    foreach ($header in $Response.Content.Headers) {
        $headers[$header.Key] = ($header.Value -join ', ')
    }

    return $headers
}

function Get-HeaderValue {
    param($Headers, [string]$Name)

    foreach ($key in $Headers.Keys) {
        if ([string]::Equals([string]$key, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [string]$Headers[$key]
        }
    }

    return $null
}

function Get-OrcaErrorDetails {
    param(
        [int]$HttpStatus,
        [string]$HttpStatusText,
        $Headers,
        [string]$RawBody
    )

    $errorType = ''
    $errorCode = ''
    $errorMessage = $RawBody
    $metadata = $null

    if (-not [string]::IsNullOrWhiteSpace($RawBody)) {
        try {
            $parsed = $RawBody | ConvertFrom-Json
            $errorObject = Get-PropertyValue -Object $parsed -Name 'error'

            if ($null -ne $errorObject) {
                $value = Get-PropertyValue -Object $errorObject -Name 'type'
                if ($null -ne $value) { $errorType = [string]$value }

                $value = Get-PropertyValue -Object $errorObject -Name 'code'
                if ($null -ne $value) { $errorCode = [string]$value }

                $value = Get-PropertyValue -Object $errorObject -Name 'message'
                if ($null -ne $value) { $errorMessage = [string]$value }

                $metadata = Get-PropertyValue -Object $errorObject -Name 'metadata'
            }
        }
        catch {
            # Keep raw body if the response is not JSON.
        }
    }

    if ([string]::IsNullOrWhiteSpace($errorMessage)) {
        $errorMessage = $HttpStatusText
    }

    $retryAfter = Get-HeaderValue -Headers $Headers -Name 'Retry-After'
    $guidance = 'error.code / error.type / HTTP Status を確認してください。'

    if ($errorCode -eq 'free_quota_exhausted') {
        $guidance = 'orcarouter/free の無料枠または利用可能な無料モデルがありません。有料モデルへ自動切替はしません。'
    }
    elseif ($HttpStatus -eq 401) {
        $guidance = 'APIキーが無効、またはAuthorizationヘッダーが不正です。'
    }
    elseif ($HttpStatus -eq 402) {
        $guidance = '支払いまたはQuotaが必要です。error.code と残高/無料枠を確認してください。'
    }
    elseif ($HttpStatus -eq 403 -and $errorCode -eq 'insufficient_user_quota') {
        $guidance = 'Workspace残高またはメンバー/エージェント予算を確認してください。'
    }
    elseif ($HttpStatus -eq 403 -and $errorCode -eq 'pre_consume_token_quota_failed') {
        $guidance = 'APIキー自身のQuota上限を確認してください。'
    }
    elseif ($HttpStatus -eq 403 -and $errorCode -eq 'access_denied') {
        $guidance = 'モデル権限・IP制限・利用上限などを確認してください。'
    }
    elseif ($HttpStatus -eq 404) {
        $guidance = 'EndpointまたはModel IDを確認してください。'
    }
    elseif ($HttpStatus -eq 425) {
        $guidance = '指定モデルはまだ利用開始前の可能性があります。error.metadataも確認してください。'
    }
    elseif ($HttpStatus -eq 429 -and -not [string]::IsNullOrWhiteSpace($retryAfter)) {
        $guidance = "Rate Limitです。Retry-After=$retryAfter 秒を待ってから再試行してください。"
    }
    elseif ($HttpStatus -eq 429) {
        $guidance = 'Rate Limitまたは無料枠制限です。時間を置いて再試行してください。'
    }
    elseif ($HttpStatus -eq 500) {
        $guidance = 'OrcaRouter内部エラーです。時間を置いて再試行してください。'
    }
    elseif ($HttpStatus -eq 502) {
        $guidance = '上流Providerまたはfallback routeが失敗しています。'
    }
    elseif ($HttpStatus -eq 503 -and $errorCode -eq 'model_not_found') {
        $guidance = '指定モデルが現在のアカウントで利用可能か確認してください。'
    }
    elseif ($HttpStatus -eq 503) {
        $guidance = 'OrcaRouterまたは上流Providerが一時的に利用できない可能性があります。'
    }

    return [pscustomobject]@{
        HttpStatus = $HttpStatus
        HttpStatusText = $HttpStatusText
        ErrorType = $errorType
        ErrorCode = $errorCode
        ErrorMessage = $errorMessage
        RetryAfter = $retryAfter
        Metadata = $metadata
        Guidance = $guidance
    }
}

function Get-AssistantText {
    param($ResponseJson)

    $choices = @(Get-PropertyValue -Object $ResponseJson -Name 'choices')

    if ($choices.Count -lt 1 -or $null -eq $choices[0]) {
        throw 'choices[0] が見つかりません。Raw responseを確認してください。'
    }

    $message = Get-PropertyValue -Object $choices[0] -Name 'message'
    if ($null -eq $message) {
        throw 'choices[0].message が見つかりません。'
    }

    $content = Get-PropertyValue -Object $message -Name 'content'

    if ($content -is [string]) {
        return $content
    }

    if ($content -is [System.Collections.IEnumerable]) {
        $parts = @()

        foreach ($part in $content) {
            if ($null -eq $part) { continue }

            $typeValue = Get-PropertyValue -Object $part -Name 'type'
            $textValue = Get-PropertyValue -Object $part -Name 'text'

            if ($typeValue -eq 'text' -and $null -ne $textValue) {
                $parts += [string]$textValue
            }
        }

        if ($parts.Count -gt 0) {
            return ($parts -join [Environment]::NewLine)
        }
    }

    throw 'choices[0].message.content を文字列として取得できません。'
}

function New-HttpClient {
    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds($requestTimeoutSeconds)
    return $client
}

function New-JsonRequest {
    param([string]$Json)

    $request = [System.Net.Http.HttpRequestMessage]::new(
        [System.Net.Http.HttpMethod]::Post,
        [string]$apiEndpoint
    )

    $request.Headers.Authorization =
        [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $ApiKey)
    [void]$request.Headers.TryAddWithoutValidation('X-OrcaRouter-Include-Cost', 'true')

    $request.Content =
        [System.Net.Http.StringContent]::new(
            $Json,
            [System.Text.Encoding]::UTF8,
            'application/json'
        )

    return $request
}

function Invoke-JsonRequest {
    param(
        $Body,
        [System.Diagnostics.Stopwatch]$Stopwatch,
        [string]$TraceTitle
    )

    $client = $null
    $request = $null
    $response = $null

    try {
        $requestJson = $Body | ConvertTo-Json -Depth 30 -Compress
        $script:lastRequest = $Body
        $script:lastResponse = $null
        $script:lastHttpStatus = $null
        $script:lastUsage = $null
        $script:lastActualModel = $Model

        Add-WorkerTrace -Step 'STEP 3' -Direction 'REQUEST' -Title $TraceTitle -Data @{
            TimeoutSeconds = $requestTimeoutSeconds
        }

        if ($PipelineSelfTest) {
            $json = [pscustomobject]@{
                id = 'mock-chat-completion'
                object = 'chat.completion'
                created = 0
                model = 'mock/provider-model'
                choices = @(
                    [pscustomobject]@{
                        index = 0
                        message = [pscustomobject]@{
                            role = 'assistant'
                            content = 'mock pipeline answer'
                        }
                        finish_reason = 'stop'
                    }
                )
                usage = [pscustomobject]@{
                    prompt_tokens = 10
                    completion_tokens = 5
                    total_tokens = 15
                }
            }

            $rawResponse = $json | ConvertTo-Json -Depth 30 -Compress
            $headers = [ordered]@{ 'X-Orca-Test' = 'pipeline-self-test' }
            $status = 200

            $script:lastHttpStatus = $status
            $script:lastResponse = $json
            $script:lastUsage = $json.usage
            $script:lastActualModel = $json.model

            Add-WorkerTrace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'Mock HTTPレスポンスを受信' -Data @{
                Status = $status
                StatusText = 'OK'
                ElapsedMs = $Stopwatch.ElapsedMilliseconds
                Headers = $headers
                RawBody = $rawResponse
            }

            return [pscustomobject]@{
                Json = $json
                RawBody = $rawResponse
                Headers = $headers
                Status = $status
            }
        }

        $client = New-HttpClient
        $request = New-JsonRequest -Json $requestJson
        $response = $client.SendAsync($request).GetAwaiter().GetResult()

        $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $headers = Get-ResponseHeaders -Response $response
        $status = [int]$response.StatusCode
        $script:lastHttpStatus = $status
        $script:lastResponse = $rawResponse

        Add-WorkerTrace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'HTTPレスポンスを受信' -Data @{
            Status = $status
            StatusText = $response.ReasonPhrase
            ElapsedMs = $Stopwatch.ElapsedMilliseconds
            Headers = $headers
            RawBody = $rawResponse
        }

        if (-not $response.IsSuccessStatusCode) {
            $details = Get-OrcaErrorDetails -HttpStatus $status -HttpStatusText $response.ReasonPhrase -Headers $headers -RawBody $rawResponse
            throw "HTTP $status : $($details.ErrorMessage) $($details.Guidance)"
        }

        try {
            $json = $rawResponse | ConvertFrom-Json
        }
        catch {
            throw "レスポンスJSONの解析に失敗しました: $($_.Exception.Message)"
        }

        $script:lastResponse = $json
        $script:lastUsage = Get-PropertyValue -Object $json -Name 'usage'
        $modelValue = Get-PropertyValue -Object $json -Name 'model'
        if (-not [string]::IsNullOrWhiteSpace([string]$modelValue)) {
            $script:lastActualModel = [string]$modelValue
        }

        return [pscustomobject]@{
            Json = $json
            RawBody = $rawResponse
            Headers = $headers
            Status = $status
        }
    }
    finally {
        if ($null -ne $response) { $response.Dispose() }
        if ($null -ne $request) { $request.Dispose() }
        if ($null -ne $client) { $client.Dispose() }
    }
}

function Invoke-Chat {
    param([System.Diagnostics.Stopwatch]$Stopwatch)

    $body = [ordered]@{
        model = $Model
        messages = @(New-ConversationMessages -CurrentQuestion $Question)
    }

    Add-WorkerTrace -Step 'STEP 2' -Direction 'REQUEST' -Title '通常Chatリクエストを組み立て' -Data @{
        Endpoint = $apiEndpoint
        Authorization = 'Bearer {0}' -f (Mask-ApiKey -Value $ApiKey)
        HistoryTurns = $historyTurnCount
        IncludeCost = $true
        Body = $body
    }

    $httpResult = Invoke-JsonRequest -Body $body -Stopwatch $Stopwatch -TraceTitle '通常Chat POSTを送信'
    $answer = Get-AssistantText -ResponseJson $httpResult.Json
    $usage = Get-PropertyValue -Object $httpResult.Json -Name 'usage'
    $actualModel = Get-PropertyValue -Object $httpResult.Json -Name 'model'

    if ([string]::IsNullOrWhiteSpace([string]$actualModel)) {
        $actualModel = $Model
    }

    Add-WorkerTrace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Assistantメッセージを解析' -Data @{
        AnswerChars = $answer.Length
        HistoryTurns = $historyTurnCount
        Usage = if ($null -ne $usage) { $usage } else { '(usage not returned)' }
    }

    return [pscustomobject]@{
        Answer = $answer
        Usage = $usage
        Request = $body
        Response = $httpResult.Json
        HttpStatus = $httpResult.Status
        ActualModel = [string]$actualModel
    }
}

function Invoke-Streaming {
    param([System.Diagnostics.Stopwatch]$Stopwatch)

    $body = [ordered]@{
        model = $Model
        messages = @(New-ConversationMessages -CurrentQuestion $Question)
        stream = $true
        stream_options = [ordered]@{ include_usage = $true }
    }

    Add-WorkerTrace -Step 'STEP 2' -Direction 'REQUEST' -Title 'Streamingリクエストを組み立て' -Data @{
        Endpoint = $apiEndpoint
        SSE = 'data: {...}, terminal data: [DONE]'
        Body = $body
    }

    $client = $null
    $request = $null
    $response = $null
    $stream = $null
    $reader = $null

    try {
        Add-WorkerTrace -Step 'STEP 3' -Direction 'REQUEST' -Title 'Streaming POSTを送信' -Data @{
            TimeoutSeconds = $requestTimeoutSeconds
        }

        $requestJson = $body | ConvertTo-Json -Depth 30 -Compress
        $client = New-HttpClient
        $request = New-JsonRequest -Json $requestJson

        $response = $client.SendAsync(
            $request,
            [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead
        ).GetAwaiter().GetResult()

        $headers = Get-ResponseHeaders -Response $response
        $status = [int]$response.StatusCode

        if (-not $response.IsSuccessStatusCode) {
            $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

            Add-WorkerTrace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'Streaming開始前にHTTPエラー' -Data @{
                Status = $status
                Headers = $headers
                RawBody = $rawResponse
            }

            $details = Get-OrcaErrorDetails -HttpStatus $status -HttpStatusText $response.ReasonPhrase -Headers $headers -RawBody $rawResponse
            throw "HTTP $status : $($details.ErrorMessage) $($details.Guidance)"
        }

        Add-WorkerTrace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'SSEストリームを開始' -Data @{
            Status = $status
            ContentType = Get-HeaderValue -Headers $headers -Name 'Content-Type'
            ElapsedMs = $Stopwatch.ElapsedMilliseconds
        }

        $stream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)

        $answer = [System.Text.StringBuilder]::new()
        $eventCount = 0
        $usage = $null
        $latestChunk = $null

        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()

            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if (-not $line.StartsWith('data:')) { continue }

            $payload = $line.Substring(5).Trim()

            if ($payload -eq '[DONE]') {
                Add-WorkerTrace -Step 'STEP 4' -Direction 'STREAM' -Title 'SSE終了 [DONE]' -Data @{
                    Events = $eventCount
                }
                continue
            }

            if ([string]::IsNullOrWhiteSpace($payload)) { continue }

            try {
                $chunk = $payload | ConvertFrom-Json
                $latestChunk = $chunk
            }
            catch {
                Add-WorkerTrace -Step 'STEP 4' -Direction 'STREAM' -Title 'JSON化できないSSE data' -Data $payload
                continue
            }

            $eventCount += 1

            if ($eventCount -le $maxStreamTraceEvents) {
                Add-WorkerTrace -Step 'STEP 4' -Direction 'STREAM' -Title "SSE data #$eventCount" -Data $chunk
            }
            elseif ($eventCount -eq ($maxStreamTraceEvents + 1)) {
                Add-WorkerTrace -Step 'STEP 4' -Direction 'STREAM' -Title 'SSE Traceを省略' -Data "可読性のため $maxStreamTraceEvents 件以降の個別イベント表示を省略します。"
            }

            $streamError = Get-PropertyValue -Object $chunk -Name 'error'

            if ($null -ne $streamError) {
                $streamMessage = Get-PropertyValue -Object $streamError -Name 'message'
                $streamType = Get-PropertyValue -Object $streamError -Name 'type'
                $streamCode = Get-PropertyValue -Object $streamError -Name 'code'
                throw "Streaming error: $streamMessage (type=$streamType, code=$streamCode)"
            }

            $choices = @(Get-PropertyValue -Object $chunk -Name 'choices')

            if ($choices.Count -gt 0 -and $null -ne $choices[0]) {
                $delta = Get-PropertyValue -Object $choices[0] -Name 'delta'

                if ($null -ne $delta) {
                    $contentPart = Get-PropertyValue -Object $delta -Name 'content'

                    if ($contentPart -is [string] -and $contentPart.Length -gt 0) {
                        [void]$answer.Append($contentPart)
                        Add-WorkerEvent -Type 'Answer' -Values @{ Text = $answer.ToString() }
                    }
                }
            }

            $usageValue = Get-PropertyValue -Object $chunk -Name 'usage'
            if ($null -ne $usageValue) { $usage = $usageValue }
        }

        Add-WorkerTrace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Streaming結果を集約' -Data @{
            AnswerChars = $answer.Length
            Usage = if ($null -ne $usage) { $usage } else { '(usage not returned)' }
        }

        $actualModel = Get-PropertyValue -Object $latestChunk -Name 'model'
        if ([string]::IsNullOrWhiteSpace([string]$actualModel)) {
            $actualModel = $Model
        }

        return [pscustomobject]@{
            Answer = $answer.ToString()
            Usage = $usage
            Request = $body
            Response = [pscustomobject]@{
                stream = $true
                latest_event = $latestChunk
                usage = $usage
                note = 'Streaming uses SSE; latest_event is the final parsed event observed by this sample.'
            }
            HttpStatus = $status
            ActualModel = [string]$actualModel
        }
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $response) { $response.Dispose() }
        if ($null -ne $request) { $request.Dispose() }
        if ($null -ne $client) { $client.Dispose() }
    }
}

function Invoke-CalculateSumTool {
    param($ArgumentsObject)

    $aValue = Get-PropertyValue -Object $ArgumentsObject -Name 'a'
    $bValue = Get-PropertyValue -Object $ArgumentsObject -Name 'b'

    if ($null -eq $aValue -or $null -eq $bValue) {
        throw 'calculate_sum の a / b が不足しています。'
    }

    $a = [double]$aValue
    $b = [double]$bValue

    return [ordered]@{
        a = $a
        b = $b
        sum = $a + $b
    }
}

function Invoke-ToolCalling {
    param([System.Diagnostics.Stopwatch]$Stopwatch)

    $tools = @(
        [ordered]@{
            type = 'function'
            function = [ordered]@{
                name = 'calculate_sum'
                description = 'Add two numbers and return the sum.'
                parameters = [ordered]@{
                    type = 'object'
                    properties = [ordered]@{
                        a = [ordered]@{ type = 'number'; description = 'First number' }
                        b = [ordered]@{ type = 'number'; description = 'Second number' }
                    }
                    required = @('a', 'b')
                    additionalProperties = $false
                }
            }
        }
    )

    $firstBody = [ordered]@{
        model = $Model
        messages = @(New-ConversationMessages -CurrentQuestion $Question)
        tools = $tools
        tool_choice = [ordered]@{
            type = 'function'
            function = [ordered]@{ name = 'calculate_sum' }
        }
    }

    Add-WorkerTrace -Step 'STEP 2' -Direction 'REQUEST' -Title 'Tool Calling 1回目のリクエストを組み立て' -Data $firstBody

    $first = Invoke-JsonRequest -Body $firstBody -Stopwatch $Stopwatch -TraceTitle 'Tool Calling 1回目を送信'

    $choices = @(Get-PropertyValue -Object $first.Json -Name 'choices')
    if ($choices.Count -lt 1 -or $null -eq $choices[0]) { throw 'choices[0] がありません。' }

    $assistantMessage = Get-PropertyValue -Object $choices[0] -Name 'message'
    if ($null -eq $assistantMessage) { throw 'choices[0].message がありません。' }

    $toolCalls = @(Get-PropertyValue -Object $assistantMessage -Name 'tool_calls')
    if ($toolCalls.Count -lt 1 -or $null -eq $toolCalls[0]) {
        throw 'tool_calls が返されませんでした。指定モデルがTool Callingに対応しているか確認してください。'
    }

    Add-WorkerTrace -Step 'STEP 5A' -Direction 'TOOL' -Title 'モデルからTool Callを受信' -Data $toolCalls

    $toolMessages = @()

    foreach ($toolCall in $toolCalls) {
        $toolCallId = [string](Get-PropertyValue -Object $toolCall -Name 'id')
        $functionObject = Get-PropertyValue -Object $toolCall -Name 'function'
        $functionName = [string](Get-PropertyValue -Object $functionObject -Name 'name')
        $argumentsJson = [string](Get-PropertyValue -Object $functionObject -Name 'arguments')

        if ($functionName -ne 'calculate_sum') {
            throw "未対応のToolが要求されました: $functionName"
        }

        try {
            $argumentsObject = $argumentsJson | ConvertFrom-Json
        }
        catch {
            throw "Tool arguments JSONを解析できません: $argumentsJson"
        }

        $toolResult = Invoke-CalculateSumTool -ArgumentsObject $argumentsObject

        Add-WorkerTrace -Step 'STEP 5B' -Direction 'TOOL' -Title 'ローカル関数 calculate_sum を実行' -Data @{
            ToolCallId = $toolCallId
            Arguments = $argumentsObject
            Result = $toolResult
        }

        $toolMessages += [ordered]@{
            role = 'tool'
            tool_call_id = $toolCallId
            content = ($toolResult | ConvertTo-Json -Depth 10 -Compress)
        }
    }

    $secondMessages = @(
        @(New-ConversationMessages -CurrentQuestion $Question)
        [ordered]@{
            role = 'assistant'
            content = Get-PropertyValue -Object $assistantMessage -Name 'content'
            tool_calls = $toolCalls
        }
    ) + $toolMessages

    $secondBody = [ordered]@{
        model = $Model
        messages = $secondMessages
    }

    Add-WorkerTrace -Step 'STEP 5C' -Direction 'REQUEST' -Title 'Tool結果を含む2回目のリクエストを組み立て' -Data $secondBody

    $second = Invoke-JsonRequest -Body $secondBody -Stopwatch $Stopwatch -TraceTitle 'Tool Calling 2回目を送信'
    $answer = Get-AssistantText -ResponseJson $second.Json
    $firstUsage = Get-PropertyValue -Object $first.Json -Name 'usage'
    $secondUsage = Get-PropertyValue -Object $second.Json -Name 'usage'
    $usage = Merge-Usage -FirstUsage $firstUsage -SecondUsage $secondUsage
    $actualModel = Get-PropertyValue -Object $second.Json -Name 'model'

    if ([string]::IsNullOrWhiteSpace([string]$actualModel)) {
        $actualModel = $Model
    }

    Add-WorkerTrace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Tool Calling後の最終回答を解析' -Data @{
        AnswerChars = $answer.Length
        HistoryTurns = $historyTurnCount
        Usage = if ($null -ne $usage) { $usage } else { '(usage not returned)' }
    }

    return [pscustomobject]@{
        Answer = $answer
        Usage = $usage
        Request = [pscustomobject]@{
            request_1 = $firstBody
            request_2 = $secondBody
        }
        Response = [pscustomobject]@{
            response_1 = $first.Json
            response_2 = $second.Json
        }
        HttpStatus = $second.Status
        ActualModel = [string]$actualModel
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Add-WorkerTrace -Step 'STEP 1' -Direction 'LOCAL' -Title '入力値を検証' -Data @{
        ApiKeyMasked = Mask-ApiKey -Value $ApiKey
        Model = $Model
        Mode = $Mode
        QuestionChars = $Question.Length
    }

    if ([string]::IsNullOrWhiteSpace($ApiKey) -or $ApiKey.StartsWith('xxx-')) {
        throw 'APIキーがダミー値のままです。OrcaRouterで発行したAPIキーを入力してください。'
    }

    if ([string]::IsNullOrWhiteSpace($Model)) { throw 'Model を入力してください。' }
    if ([string]::IsNullOrWhiteSpace($Question)) { throw '質問を入力してください。' }

    if ($Mode -eq 'Streaming') {
        $result = Invoke-Streaming -Stopwatch $stopwatch
    }
    elseif ($Mode -eq 'Tool Calling') {
        $result = Invoke-ToolCalling -Stopwatch $stopwatch
    }
    else {
        $result = Invoke-Chat -Stopwatch $stopwatch
    }

    Add-WorkerTrace -Step 'STEP 6' -Direction 'LOCAL' -Title 'バックグラウンド処理を完了' -Data @{
        Mode = $Mode
        Completed = $true
        HistoryTurns = $historyTurnCount
        TotalElapsedMs = $stopwatch.ElapsedMilliseconds
    }

    Add-WorkerEvent -Type 'Completed' -Values @{
        Answer = [string]$result.Answer
        TotalElapsedMs = $stopwatch.ElapsedMilliseconds
        HttpStatus = $result.HttpStatus
        Usage = $result.Usage
        Request = $result.Request
        Response = $result.Response
        ActualModel = [string]$result.ActualModel
    }
}
catch {
    Add-WorkerEvent -Type 'Error' -Values @{
        Message = $_.Exception.Message
        ExceptionType = $_.Exception.GetType().FullName
        ScriptLineNumber = $_.InvocationInfo.ScriptLineNumber
        PositionMessage = $_.InvocationInfo.PositionMessage
        ScriptStackTrace = $_.ScriptStackTrace
        TotalElapsedMs = $stopwatch.ElapsedMilliseconds
        HttpStatus = $script:lastHttpStatus
        Usage = $script:lastUsage
        Request = $script:lastRequest
        Response = $script:lastResponse
        ActualModel = $script:lastActualModel
    }
}
finally {
    $stopwatch.Stop()
}
