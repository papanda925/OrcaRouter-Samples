#requires -Version 5.1
<#
.SYNOPSIS
    OrcaRouter API learning sample using PowerShell + WPF/XAML.

.DESCRIPTION
    Chat / Streaming / Tool Calling are implemented with the same learning flow:

    STEP 1 - Validate inputs
    STEP 2 - Build request
    STEP 3 - Send HTTP POST
    STEP 4 - Receive response
    STEP 5 - Parse / process result
    STEP 6 - Update UI and trace

    Streaming follows the OpenAI-compatible SSE format.
    Tool Calling demonstrates a local calculate_sum function.

    A real API key is never written to the trace.
#>

param(
    [switch]$SyntaxCheck
)

if ($SyntaxCheck) {
    exit 0
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Net.Http

if (-not ('OrcaRouterViewModel' -as [type])) {
    Add-Type -TypeDefinition @'
using System.ComponentModel;

public sealed class OrcaRouterViewModel : INotifyPropertyChanged
{
    private string _model = "";
    private string _mode = "";
    private string _question = "";
    private string _answer = "";
    private string _statusText = "";

    public event PropertyChangedEventHandler PropertyChanged;

    private void RaisePropertyChanged(string propertyName)
    {
        var handler = PropertyChanged;
        if (handler != null)
        {
            handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public string Model
    {
        get { return _model; }
        set
        {
            if (_model == value) return;
            _model = value;
            RaisePropertyChanged("Model");
        }
    }

    public string Mode
    {
        get { return _mode; }
        set
        {
            if (_mode == value) return;
            _mode = value;
            RaisePropertyChanged("Mode");
        }
    }

    public string Question
    {
        get { return _question; }
        set
        {
            if (_question == value) return;
            _question = value;
            RaisePropertyChanged("Question");
        }
    }

    public string Answer
    {
        get { return _answer; }
        set
        {
            if (_answer == value) return;
            _answer = value;
            RaisePropertyChanged("Answer");
        }
    }

    public string StatusText
    {
        get { return _statusText; }
        set
        {
            if (_statusText == value) return;
            _statusText = value;
            RaisePropertyChanged("StatusText");
        }
    }
}
'@
}

$script:ApiEndpoint = 'https://api.orcarouter.ai/v1/chat/completions'
$script:ApiKeyPlaceholder = 'xxx-your-orcarouter-api-key-xxx'

# LOCAL TEST ONLY:
# 実APIキーをソースへ一時的に埋め込んで試す場合は、次の値だけを書き換えます。
# 公開GitHubへ実APIキーをコミットしないでください。
$script:DefaultApiKey = 'xxx-your-orcarouter-api-key-xxx'

$script:RequestTimeoutSeconds = 60
$script:MaxStreamTraceEvents = 50

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'WPF requires an STA thread. Start PowerShell with -STA and run this script again.'
}

function Mask-ApiKey {
    param([string]$ApiKey)

    if ([string]::IsNullOrWhiteSpace($ApiKey)) { return '(empty)' }
    if ($ApiKey.Length -le 8) { return '********' }

    return '{0}...{1}' -f $ApiKey.Substring(0, 4), $ApiKey.Substring($ApiKey.Length - 4)
}

function ConvertTo-TraceText {
    param($Data)

    if ($null -eq $Data) { return '' }
    if ($Data -is [string]) { return $Data }

    try {
        return ($Data | ConvertTo-Json -Depth 30)
    }
    catch {
        return [string]$Data
    }
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
            # Keep raw body when the response is not JSON.
        }
    }

    if ([string]::IsNullOrWhiteSpace($errorMessage)) {
        $errorMessage = $HttpStatusText
    }

    $retryAfter = Get-HeaderValue -Headers $Headers -Name 'Retry-After'
    $guidance = 'error.code / error.type / HTTP Status をTraceで確認してください。'

    if ($HttpStatus -eq 401) {
        $guidance = 'APIキーが無効、またはAuthorizationヘッダーが不正です。'
    }
    elseif ($HttpStatus -eq 403 -and $errorCode -eq 'insufficient_user_quota') {
        $guidance = 'Workspace残高またはメンバー/エージェント予算を確認してください。'
    }
    elseif ($HttpStatus -eq 403 -and $errorCode -eq 'pre_consume_token_quota_failed') {
        $guidance = 'APIキー自身のQuota上限を確認してください。'
    }
    elseif ($HttpStatus -eq 403 -and $errorCode -eq 'free_quota_exhausted') {
        $guidance = '無料ルーターで処理可能なモデルがありません。有料モデルを指定してください。'
    }
    elseif ($HttpStatus -eq 429 -and -not [string]::IsNullOrWhiteSpace($retryAfter)) {
        $guidance = "Rate Limitです。Retry-After=$retryAfter 秒を待ってから再試行してください。"
    }
    elseif ($HttpStatus -eq 429) {
        $guidance = 'Retry-Afterのない無料枠429は、同じPromptを待って再送しても改善しない場合があります。'
    }
    elseif ($HttpStatus -eq 425) {
        $guidance = '指定モデルはまだ利用開始前の可能性があります。error.metadataも確認してください。'
    }
    elseif ($HttpStatus -eq 503 -and $errorCode -eq 'model_not_found') {
        $guidance = '指定モデルが現在のアカウントで利用可能か確認してください。'
    }

    return [pscustomobject]@{
        HttpStatus     = $HttpStatus
        HttpStatusText = $HttpStatusText
        ErrorType      = $errorType
        ErrorCode      = $errorCode
        ErrorMessage   = $errorMessage
        RetryAfter     = $retryAfter
        Metadata       = $metadata
        Guidance       = $guidance
    }
}

function Get-AssistantText {
    param($ResponseJson)

    $choices = @(Get-PropertyValue -Object $ResponseJson -Name 'choices')

    if ($choices.Count -lt 1 -or $null -eq $choices[0]) {
        throw 'choices[0] が見つかりません。Trace の Raw response を確認してください。'
    }

    $message = Get-PropertyValue -Object $choices[0] -Name 'message'
    if ($null -eq $message) {
        throw 'choices[0].message が見つかりません。Trace の Raw response を確認してください。'
    }

    $content = Get-PropertyValue -Object $message -Name 'content'

    if ($content -is [string]) { return $content }

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

    throw 'choices[0].message.content を文字列として取得できません。Trace の Raw response を確認してください。'
}

$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
if (-not (Test-Path -LiteralPath $xamlPath)) {
    throw 'MainWindow.xaml was not found next to OrcaRouterChat.ps1.'
}

[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw -Encoding UTF8
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$apiKeyBox = $window.FindName('ApiKeyBox')
$modelBox = $window.FindName('ModelBox')
$modeBox = $window.FindName('ModeBox')
$questionBox = $window.FindName('QuestionBox')
$answerBox = $window.FindName('AnswerBox')
$traceBox = $window.FindName('TraceBox')
$sendButton = $window.FindName('SendButton')
$clearTraceButton = $window.FindName('ClearTraceButton')
$statusText = $window.FindName('StatusText')

$viewModel = [OrcaRouterViewModel]::new()
$viewModel.Model = 'orcarouter/free'
$viewModel.Mode = 'Chat'
$viewModel.Question = '日本語で「こんにちは。PowerShell版Chatのテストです。」とだけ答えてください。'
$viewModel.Answer = 'ここに回答が表示されます。'
$viewModel.StatusText = 'Ready'
$window.DataContext = $viewModel

$apiKeyBox.Password = $script:DefaultApiKey

function Add-Trace {
    param(
        [string]$Step,
        [string]$Direction,
        [string]$Title,
        $Data
    )

    $timestamp = Get-Date -Format 'HH:mm:ss.fff'
    $separator = '-' * 88
    $dataText = ConvertTo-TraceText -Data $Data

    $block = @(
        $separator
        "[$timestamp] [$Step] [$Direction] $Title"
    )

    if (-not [string]::IsNullOrWhiteSpace($dataText)) {
        $block += $dataText
    }

    $block += ''

    $traceBox.AppendText(($block -join [Environment]::NewLine))
    $traceBox.ScrollToEnd()
}

function Set-UiBusy {
    param([bool]$Busy)

    $sendButton.IsEnabled = -not $Busy
    $apiKeyBox.IsEnabled = -not $Busy
    $modelBox.IsEnabled = -not $Busy
    $modeBox.IsEnabled = -not $Busy

    # Keep Question editable even while an API request is running.
    # This also avoids making the input box look broken during a slow model call.
    $questionBox.IsEnabled = $true
}

function Refresh-Ui {
    $window.Dispatcher.Invoke(
        [System.Action]{ },
        [System.Windows.Threading.DispatcherPriority]::Background
    )
}

function Get-SelectedMode {
    if ([string]::IsNullOrWhiteSpace($viewModel.Mode)) {
        return 'Chat'
    }

    return [string]$viewModel.Mode
}

function Assert-Inputs {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [string]$Mode
    )

    Add-Trace -Step 'STEP 1' -Direction 'LOCAL' -Title '入力値を検証' -Data ([ordered]@{
        ApiKeyMasked = Mask-ApiKey -ApiKey $ApiKey
        Model = $Model
        Mode = $Mode
        QuestionChars = $Question.Length
    })

    if ([string]::IsNullOrWhiteSpace($ApiKey) -or
        $ApiKey -eq $script:ApiKeyPlaceholder -or
        $ApiKey.StartsWith('xxx-')) {
        throw 'APIキーがダミー値のままです。OrcaRouterで発行したAPIキーを入力してください。'
    }

    if ([string]::IsNullOrWhiteSpace($Model)) { throw 'Model を入力してください。' }
    if ([string]::IsNullOrWhiteSpace($Question)) { throw '質問を入力してください。' }
}

function New-HttpClient {
    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds($script:RequestTimeoutSeconds)
    return $client
}

function New-JsonRequest {
    param([string]$ApiKey, [string]$Json)

    $request = [System.Net.Http.HttpRequestMessage]::new(
        [System.Net.Http.HttpMethod]::Post,
        [string]$script:ApiEndpoint
    )

    $request.Headers.Authorization =
        [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $ApiKey)

    $request.Content =
        [System.Net.Http.StringContent]::new(
            $Json,
            [System.Text.Encoding]::UTF8,
            'application/json'
        )

    return $request
}

function Invoke-OrcaJsonRequest {
    param(
        [string]$ApiKey,
        $Body,
        [System.Diagnostics.Stopwatch]$Stopwatch,
        [string]$TraceTitle
    )

    $client = $null
    $request = $null
    $response = $null

    try {
        $requestJson = $Body | ConvertTo-Json -Depth 30 -Compress

        Add-Trace -Step 'STEP 3' -Direction 'REQUEST' -Title $TraceTitle -Data ([ordered]@{
            TimeoutSeconds = $script:RequestTimeoutSeconds
        })

        $client = New-HttpClient
        $request = New-JsonRequest -ApiKey $ApiKey -Json $requestJson
        $response = $client.SendAsync($request).GetAwaiter().GetResult()

        $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $headers = Get-ResponseHeaders -Response $response
        $status = [int]$response.StatusCode

        Add-Trace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'HTTPレスポンスを受信' -Data ([ordered]@{
            Status = $status
            StatusText = $response.ReasonPhrase
            ElapsedMs = $Stopwatch.ElapsedMilliseconds
            Headers = $headers
            RawBody = $rawResponse
        })

        if (-not $response.IsSuccessStatusCode) {
            $details = Get-OrcaErrorDetails -HttpStatus $status -HttpStatusText $response.ReasonPhrase -Headers $headers -RawBody $rawResponse
            $exception = [System.Exception]::new("HTTP ${status}: $($details.ErrorMessage) $($details.Guidance)")
            $exception.Data['OrcaErrorDetails'] = (ConvertTo-TraceText $details)
            throw $exception
        }

        try {
            $json = $rawResponse | ConvertFrom-Json
        }
        catch {
            throw "レスポンスJSONの解析に失敗しました: $($_.Exception.Message)"
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

function Invoke-ChatMode {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [System.Diagnostics.Stopwatch]$Stopwatch
    )

    $body = [ordered]@{
        model = $Model
        messages = @(
            [ordered]@{ role = 'user'; content = $Question }
        )
    }

    Add-Trace -Step 'STEP 2' -Direction 'REQUEST' -Title '通常Chatリクエストを組み立て' -Data ([ordered]@{
        Endpoint = $script:ApiEndpoint
        Authorization = 'Bearer {0}' -f (Mask-ApiKey -ApiKey $ApiKey)
        Body = $body
    })

    $result = Invoke-OrcaJsonRequest -ApiKey $ApiKey -Body $body -Stopwatch $Stopwatch -TraceTitle '通常Chat POSTを送信'
    $assistantText = Get-AssistantText -ResponseJson $result.Json
    $usage = Get-PropertyValue -Object $result.Json -Name 'usage'

    Add-Trace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Assistantメッセージを解析' -Data ([ordered]@{
        AnswerChars = $assistantText.Length
        Usage = if ($null -ne $usage) { $usage } else { '(usage not returned)' }
    })

    return $assistantText
}

function Invoke-StreamingMode {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [System.Diagnostics.Stopwatch]$Stopwatch
    )

    $body = [ordered]@{
        model = $Model
        messages = @(
            [ordered]@{ role = 'user'; content = $Question }
        )
        stream = $true
        stream_options = [ordered]@{ include_usage = $true }
    }

    Add-Trace -Step 'STEP 2' -Direction 'REQUEST' -Title 'Streamingリクエストを組み立て' -Data ([ordered]@{
        Endpoint = $script:ApiEndpoint
        Body = $body
        SSE = 'data: {...}, terminal data: [DONE]'
    })

    $client = $null
    $request = $null
    $response = $null
    $stream = $null
    $streamReader = $null

    try {
        Add-Trace -Step 'STEP 3' -Direction 'REQUEST' -Title 'Streaming POSTを送信' -Data ([ordered]@{
            TimeoutSeconds = $script:RequestTimeoutSeconds
        })

        $requestJson = $body | ConvertTo-Json -Depth 30 -Compress
        $client = New-HttpClient
        $request = New-JsonRequest -ApiKey $ApiKey -Json $requestJson

        $response = $client.SendAsync(
            $request,
            [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead
        ).GetAwaiter().GetResult()

        $headers = Get-ResponseHeaders -Response $response
        $status = [int]$response.StatusCode

        if (-not $response.IsSuccessStatusCode) {
            $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

            Add-Trace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'Streaming開始前にHTTPエラー' -Data ([ordered]@{
                Status = $status
                Headers = $headers
                RawBody = $rawResponse
            })

            $details = Get-OrcaErrorDetails -HttpStatus $status -HttpStatusText $response.ReasonPhrase -Headers $headers -RawBody $rawResponse
            throw "HTTP ${status}: $($details.ErrorMessage) $($details.Guidance)"
        }

        Add-Trace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'SSEストリームを開始' -Data ([ordered]@{
            Status = $status
            ContentType = Get-HeaderValue -Headers $headers -Name 'Content-Type'
            ElapsedMs = $Stopwatch.ElapsedMilliseconds
        })

        $stream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $streamReader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)

        $answer = [System.Text.StringBuilder]::new()
        $eventCount = 0
        $usage = $null

        while (-not $streamReader.EndOfStream) {
            $line = $streamReader.ReadLine()

            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if (-not $line.StartsWith('data:')) { continue }

            $payload = $line.Substring(5).Trim()

            if ($payload -eq '[DONE]') {
                Add-Trace -Step 'STEP 4' -Direction 'STREAM' -Title 'SSE終了 [DONE]' -Data ([ordered]@{
                    Events = $eventCount
                })
                continue
            }

            if ([string]::IsNullOrWhiteSpace($payload)) { continue }

            try {
                $chunk = $payload | ConvertFrom-Json
            }
            catch {
                Add-Trace -Step 'STEP 4' -Direction 'STREAM' -Title 'JSON化できないSSE data' -Data $payload
                continue
            }

            $eventCount += 1

            if ($eventCount -le $script:MaxStreamTraceEvents) {
                Add-Trace -Step 'STEP 4' -Direction 'STREAM' -Title "SSE data #$eventCount" -Data $chunk
            }
            elseif ($eventCount -eq ($script:MaxStreamTraceEvents + 1)) {
                Add-Trace -Step 'STEP 4' -Direction 'STREAM' -Title 'SSE Traceを省略' -Data "可読性のため $($script:MaxStreamTraceEvents) 件以降の個別イベント表示を省略します。"
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
                        $viewModel.Answer = $answer.ToString()
                        $answerBox.ScrollToEnd()
                        Refresh-Ui
                    }
                }
            }

            $usageValue = Get-PropertyValue -Object $chunk -Name 'usage'
            if ($null -ne $usageValue) { $usage = $usageValue }
        }

        Add-Trace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Streaming結果を集約' -Data ([ordered]@{
            AnswerChars = $answer.Length
            Usage = if ($null -ne $usage) { $usage } else { '(usage not returned)' }
        })

        return $answer.ToString()
    }
    finally {
        if ($null -ne $streamReader) { $streamReader.Dispose() }
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

function Invoke-ToolCallingMode {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [System.Diagnostics.Stopwatch]$Stopwatch
    )

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
        messages = @(
            [ordered]@{ role = 'user'; content = $Question }
        )
        tools = $tools
        tool_choice = [ordered]@{
            type = 'function'
            function = [ordered]@{ name = 'calculate_sum' }
        }
    }

    Add-Trace -Step 'STEP 2' -Direction 'REQUEST' -Title 'Tool Calling 1回目のリクエストを組み立て' -Data $firstBody

    $first = Invoke-OrcaJsonRequest -ApiKey $ApiKey -Body $firstBody -Stopwatch $Stopwatch -TraceTitle 'Tool Calling 1回目を送信'

    $choices = @(Get-PropertyValue -Object $first.Json -Name 'choices')
    if ($choices.Count -lt 1 -or $null -eq $choices[0]) { throw 'choices[0] がありません。' }

    $assistantMessage = Get-PropertyValue -Object $choices[0] -Name 'message'
    if ($null -eq $assistantMessage) { throw 'choices[0].message がありません。' }

    $toolCalls = @(Get-PropertyValue -Object $assistantMessage -Name 'tool_calls')
    if ($toolCalls.Count -lt 1 -or $null -eq $toolCalls[0]) {
        throw 'tool_calls が返されませんでした。指定モデルがTool Callingに対応しているか確認してください。'
    }

    Add-Trace -Step 'STEP 5A' -Direction 'TOOL' -Title 'モデルからTool Callを受信' -Data $toolCalls

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

        Add-Trace -Step 'STEP 5B' -Direction 'TOOL' -Title 'ローカル関数 calculate_sum を実行' -Data ([ordered]@{
            ToolCallId = $toolCallId
            Arguments = $argumentsObject
            Result = $toolResult
        })

        $toolMessages += [ordered]@{
            role = 'tool'
            tool_call_id = $toolCallId
            content = ($toolResult | ConvertTo-Json -Depth 10 -Compress)
        }
    }

    $secondMessages = @(
        [ordered]@{ role = 'user'; content = $Question },
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

    Add-Trace -Step 'STEP 5C' -Direction 'REQUEST' -Title 'Tool結果を含む2回目のリクエストを組み立て' -Data $secondBody

    $second = Invoke-OrcaJsonRequest -ApiKey $ApiKey -Body $secondBody -Stopwatch $Stopwatch -TraceTitle 'Tool Calling 2回目を送信'

    $assistantText = Get-AssistantText -ResponseJson $second.Json
    $usage = Get-PropertyValue -Object $second.Json -Name 'usage'

    Add-Trace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Tool Calling後の最終回答を解析' -Data ([ordered]@{
        AnswerChars = $assistantText.Length
        Usage = if ($null -ne $usage) { $usage } else { '(usage not returned)' }
    })

    return $assistantText
}

function Protect-LocalTraceText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    $sanitized = $Text

    # Keep traces useful without exposing a local user name or repository path.
    # Replace the most common machine-specific path roots before displaying an
    # exception/stack trace in the sample UI.
    $replacements = @(
        [pscustomobject]@{
            Value = $PSScriptRoot
            Placeholder = '<repository-root>\PowerShell'
        },
        [pscustomobject]@{
            Value = [Environment]::GetFolderPath('UserProfile')
            Placeholder = '<USERPROFILE>'
        },
        [pscustomobject]@{
            Value = $env:HOME
            Placeholder = '<HOME>'
        }
    )

    foreach ($replacement in $replacements) {
        if (-not [string]::IsNullOrWhiteSpace($replacement.Value)) {
            $sanitized = $sanitized.Replace(
                [string]$replacement.Value,
                [string]$replacement.Placeholder
            )
        }
    }

    return $sanitized
}

function Invoke-OrcaRouterChat {
    $apiKey = $apiKeyBox.Password.Trim()
    $model = $viewModel.Model.Trim()
    $question = $viewModel.Question.Trim()
    $mode = Get-SelectedMode

    $viewModel.Answer = ''
    $viewModel.StatusText = 'Processing...'
    $statusText.Foreground = '#64748B'
    Set-UiBusy -Busy $true

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        Assert-Inputs -ApiKey $apiKey -Model $model -Question $question -Mode $mode

        if ($mode -eq 'Streaming') {
            $assistantText = Invoke-StreamingMode -ApiKey $apiKey -Model $model -Question $question -Stopwatch $stopwatch
        }
        elseif ($mode -eq 'Tool Calling') {
            $assistantText = Invoke-ToolCallingMode -ApiKey $apiKey -Model $model -Question $question -Stopwatch $stopwatch
        }
        else {
            $assistantText = Invoke-ChatMode -ApiKey $apiKey -Model $model -Question $question -Stopwatch $stopwatch
        }

        # STEP 6: Update UI and trace.
        $viewModel.Answer = $assistantText

        Add-Trace -Step 'STEP 6' -Direction 'LOCAL' -Title '画面へ回答を表示' -Data ([ordered]@{
            Mode = $mode
            Completed = $true
            TotalElapsedMs = $stopwatch.ElapsedMilliseconds
        })

        $viewModel.StatusText = 'Completed'
        $statusText.Foreground = '#0F766E'
    }
    catch {
        $exception = $_.Exception

        $details = $null
        if ($exception.Data.Contains('OrcaErrorDetails')) {
            $details = $exception.Data['OrcaErrorDetails']
        }

        $safeMessage = Protect-LocalTraceText -Text $exception.Message
        $safeScriptStack = Protect-LocalTraceText -Text $_.ScriptStackTrace
        $safePosition = Protect-LocalTraceText -Text $_.InvocationInfo.PositionMessage

        Add-Trace -Step 'ERROR' -Direction 'ERROR' -Title '処理中にエラーが発生' -Data ([ordered]@{
            Message = $safeMessage
            Type = $exception.GetType().FullName
            OrcaErrorDetails = $details
            ScriptStack = $safeScriptStack
            Position = $safePosition
            Privacy = 'Local user/repository path roots are replaced with placeholders.'
        })

        $viewModel.Answer = "ERROR: $safeMessage"
        $viewModel.StatusText = 'Error - Trace を確認してください'
        $statusText.Foreground = '#B42318'
    }
    finally {
        $stopwatch.Stop()
        Set-UiBusy -Busy $false
    }
}

$sendButton.Add_Click({
    Invoke-OrcaRouterChat
})

$clearTraceButton.Add_Click({
    $traceBox.Clear()
    $viewModel.StatusText = 'Ready'
    $statusText.Foreground = '#64748B'
})

$modeBox.Add_SelectionChanged({
    $mode = Get-SelectedMode

    if ($mode -eq 'Tool Calling') {
        $viewModel.Question = 'calculate_sum ツールを使って 123 と 456 を足し、その結果を日本語で説明してください。'
    }
    elseif ($mode -eq 'Streaming') {
        $viewModel.Question = '日本語で「こんにちは。Streamingのテストです。」と短く答えてください。'
    }
    elseif ($mode -eq 'Chat') {
        $viewModel.Question = '日本語で「こんにちは。PowerShell版Chatのテストです。」とだけ答えてください。'
    }
})

$questionBox.Add_PreviewKeyDown({
    param($sender, $eventArgs)

    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Enter -and
        ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
        $eventArgs.Handled = $true
        Invoke-OrcaRouterChat
    }
})

Add-Trace -Step 'READY' -Direction 'LOCAL' -Title 'サンプルを起動' -Data ([ordered]@{
    Endpoint = $script:ApiEndpoint
    Model = $viewModel.Model
    Mode = Get-SelectedMode
    Note = 'APIキーはダミー値です。実行前に画面上で差し替えてください。'
})

$window.Add_ContentRendered({
    # Explicitly focus the editable question box and keep IME/input enabled.
    [void]$questionBox.Focus()
    $questionBox.CaretIndex = $questionBox.Text.Length
})

$null = $window.ShowDialog()
