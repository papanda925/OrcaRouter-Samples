#requires -Version 5.1
<#
.SYNOPSIS
    OrcaRouter API learning sample using PowerShell + WPF/XAML.

.DESCRIPTION
    The Web, PowerShell, and VBA samples intentionally follow the same
    six processing steps so learners can compare the implementations.

    STEP 1 - Validate inputs
    STEP 2 - Build request
    STEP 3 - Send HTTP POST
    STEP 4 - Receive response
    STEP 5 - Parse assistant message
    STEP 6 - Update UI and trace

    A real API key is never written to the trace.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Net.Http

$script:ApiEndpoint = 'https://api.orcarouter.ai/v1/chat/completions'
$script:ApiKeyPlaceholder = 'xxx-your-orcarouter-api-key-xxx'

# LOCAL TEST ONLY:
# 実APIキーをソースへ一時的に埋め込んで試す場合は、次の値だけを書き換えます。
# 公開GitHubへ実APIキーをコミットしないでください。
$script:DefaultApiKey = 'xxx-your-orcarouter-api-key-xxx'

$script:RequestTimeoutSeconds = 60

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'WPF requires an STA thread. Start PowerShell with -STA and run this script again.'
}

function Mask-ApiKey {
    param(
        [Parameter(Mandatory = $false)]
        [string]$ApiKey
    )

    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return '(empty)'
    }

    if ($ApiKey.Length -le 8) {
        return '********'
    }

    return '{0}...{1}' -f $ApiKey.Substring(0, 4), $ApiKey.Substring($ApiKey.Length - 4)
}

function ConvertTo-TraceText {
    param(
        [Parameter(Mandatory = $false)]
        $Data
    )

    if ($null -eq $Data) {
        return ''
    }

    if ($Data -is [string]) {
        return $Data
    }

    try {
        return ($Data | ConvertTo-Json -Depth 20)
    }
    catch {
        return [string]$Data
    }
}

function Get-ResponseHeaders {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.Http.HttpResponseMessage]$Response
    )

    $headers = [ordered]@{}

    foreach ($header in $Response.Headers) {
        $headers[$header.Key] = ($header.Value -join ', ')
    }

    foreach ($header in $Response.Content.Headers) {
        $headers[$header.Key] = ($header.Value -join ', ')
    }

    return $headers
}

function Get-AssistantText {
    param(
        [Parameter(Mandatory = $true)]
        $ResponseJson
    )

    $choicesProperty = $ResponseJson.PSObject.Properties['choices']

    if ($null -eq $choicesProperty) {
        throw 'choices が見つかりません。Trace の Raw response を確認してください。'
    }

    $choices = @($choicesProperty.Value)

    if ($choices.Count -lt 1) {
        throw 'choices[0] が見つかりません。Trace の Raw response を確認してください。'
    }

    $messageProperty = $choices[0].PSObject.Properties['message']

    if ($null -eq $messageProperty) {
        throw 'choices[0].message が見つかりません。Trace の Raw response を確認してください。'
    }

    $contentProperty = $messageProperty.Value.PSObject.Properties['content']

    if ($null -eq $contentProperty) {
        throw 'choices[0].message.content が見つかりません。Trace の Raw response を確認してください。'
    }

    $content = $contentProperty.Value

    if ($content -is [string]) {
        return $content
    }

    if ($content -is [System.Collections.IEnumerable]) {
        $parts = @()

        foreach ($part in $content) {
            if ($null -eq $part) {
                continue
            }

            $typeProperty = $part.PSObject.Properties['type']
            $textProperty = $part.PSObject.Properties['text']

            if ($null -ne $typeProperty -and
                $null -ne $textProperty -and
                $typeProperty.Value -eq 'text' -and
                $textProperty.Value) {
                $parts += [string]$textProperty.Value
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
    throw "XAML file was not found: $xamlPath"
}

[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw -Encoding UTF8
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$apiKeyBox = $window.FindName('ApiKeyBox')
$modelBox = $window.FindName('ModelBox')
$questionBox = $window.FindName('QuestionBox')
$answerBox = $window.FindName('AnswerBox')
$traceBox = $window.FindName('TraceBox')
$sendButton = $window.FindName('SendButton')
$clearTraceButton = $window.FindName('ClearTraceButton')
$statusText = $window.FindName('StatusText')

$apiKeyBox.Password = $script:DefaultApiKey

function Add-Trace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Step,

        [Parameter(Mandatory = $true)]
        [string]$Direction,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
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
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Busy
    )

    $sendButton.IsEnabled = -not $Busy
    $apiKeyBox.IsEnabled = -not $Busy
    $modelBox.IsEnabled = -not $Busy
    $questionBox.IsEnabled = -not $Busy
}

function Invoke-OrcaRouterChat {
    $apiKey = $apiKeyBox.Password.Trim()
    $model = $modelBox.Text.Trim()
    $question = $questionBox.Text.Trim()

    $answerBox.Text = ''
    $statusText.Text = 'Processing...'
    $statusText.Foreground = '#64748B'
    Set-UiBusy -Busy $true

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $client = $null
    $request = $null
    $response = $null
    $rawResponse = ''
    $httpStatus = '(no HTTP response)'

    try {
        # STEP 1: Validate inputs.
        Add-Trace -Step 'STEP 1' -Direction 'LOCAL' -Title '入力値を検証' -Data ([ordered]@{
            ApiKeyMasked  = Mask-ApiKey -ApiKey $apiKey
            Model         = $model
            QuestionChars = $question.Length
        })

        if ([string]::IsNullOrWhiteSpace($apiKey) -or
            $apiKey -eq $script:ApiKeyPlaceholder -or
            $apiKey.StartsWith('xxx-')) {
            throw 'APIキーがダミー値のままです。OrcaRouterで発行したAPIキーを入力してください。'
        }

        if ([string]::IsNullOrWhiteSpace($model)) {
            throw 'Model を入力してください。'
        }

        if ([string]::IsNullOrWhiteSpace($question)) {
            throw '質問を入力してください。'
        }

        # STEP 2: Build request.
        $requestObject = [ordered]@{
            model = $model
            messages = @(
                [ordered]@{
                    role = 'user'
                    content = $question
                }
            )
        }

        $requestJson = $requestObject | ConvertTo-Json -Depth 10 -Compress

        Add-Trace -Step 'STEP 2' -Direction 'REQUEST' -Title 'HTTPリクエストを組み立て' -Data ([ordered]@{
            Method   = 'POST'
            Endpoint = $script:ApiEndpoint
            Headers  = [ordered]@{
                Authorization = 'Bearer {0}' -f (Mask-ApiKey -ApiKey $apiKey)
                'Content-Type' = 'application/json'
            }
            Body = $requestObject
        })

        # STEP 3: Send HTTP POST.
        Add-Trace -Step 'STEP 3' -Direction 'REQUEST' -Title 'OrcaRouterへPOSTを送信' -Data ([ordered]@{
            TimeoutSeconds = $script:RequestTimeoutSeconds
        })

        $client = [System.Net.Http.HttpClient]::new()
        $client.Timeout = [TimeSpan]::FromSeconds($script:RequestTimeoutSeconds)

        $request = [System.Net.Http.HttpRequestMessage]::new(
            [System.Net.Http.HttpMethod]::Post,
            [string]$script:ApiEndpoint
        )

        $request.Headers.Authorization =
            [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $apiKey)

        $request.Content =
            [System.Net.Http.StringContent]::new(
                $requestJson,
                [System.Text.Encoding]::UTF8,
                'application/json'
            )

        $response = $client.SendAsync($request).GetAwaiter().GetResult()

        # STEP 4: Receive response.
        $httpStatus = '{0} {1}' -f [int]$response.StatusCode, $response.ReasonPhrase
        $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $responseHeaders = Get-ResponseHeaders -Response $response

        Add-Trace -Step 'STEP 4' -Direction 'RESPONSE' -Title 'HTTPレスポンスを受信' -Data ([ordered]@{
            Status    = $httpStatus
            ElapsedMs = $stopwatch.ElapsedMilliseconds
            Headers   = $responseHeaders
            RawBody   = $rawResponse
        })

        if (-not $response.IsSuccessStatusCode) {
            throw ('HTTP request failed: {0}{1}{2}' -f $httpStatus, [Environment]::NewLine, $rawResponse)
        }

        # STEP 5: Parse assistant message.
        try {
            $responseJson = $rawResponse | ConvertFrom-Json
        }
        catch {
            throw "レスポンスJSONの解析に失敗しました: $($_.Exception.Message)"
        }

        $assistantText = Get-AssistantText -ResponseJson $responseJson

        $usageProperty = $responseJson.PSObject.Properties['usage']

        Add-Trace -Step 'STEP 5' -Direction 'LOCAL' -Title 'Assistantメッセージを解析' -Data ([ordered]@{
            AnswerChars = $assistantText.Length
            Usage       = if ($null -ne $usageProperty) { $usageProperty.Value } else { '(usage not returned)' }
        })

        # STEP 6: Update UI and trace.
        $answerBox.Text = $assistantText

        Add-Trace -Step 'STEP 6' -Direction 'LOCAL' -Title '画面へ回答を表示' -Data ([ordered]@{
            Completed      = $true
            TotalElapsedMs = $stopwatch.ElapsedMilliseconds
        })

        $statusText.Text = 'Completed'
        $statusText.Foreground = '#0F766E'
    }
    catch {
        $exception = $_.Exception

        Add-Trace -Step 'ERROR' -Direction 'ERROR' -Title '処理中にエラーが発生' -Data ([ordered]@{
            Message      = $exception.Message
            Type         = $exception.GetType().FullName
            HttpStatus   = $httpStatus
            RawResponse  = if ($rawResponse) { $rawResponse } else { '(not available)' }
            ScriptStack  = $_.ScriptStackTrace
            Position     = $_.InvocationInfo.PositionMessage
        })

        $answerBox.Text = "ERROR: $($exception.Message)"
        $statusText.Text = 'Error - Trace を確認してください'
        $statusText.Foreground = '#B42318'
    }
    finally {
        $stopwatch.Stop()

        if ($null -ne $response) {
            $response.Dispose()
        }

        if ($null -ne $request) {
            $request.Dispose()
        }

        if ($null -ne $client) {
            $client.Dispose()
        }

        Set-UiBusy -Busy $false
    }
}

$sendButton.Add_Click({
    Invoke-OrcaRouterChat
})

$clearTraceButton.Add_Click({
    $traceBox.Clear()
    $statusText.Text = 'Ready'
    $statusText.Foreground = '#64748B'
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
    Model    = $modelBox.Text
    Note     = 'APIキーはダミー値です。実行前に画面上で差し替えてください。'
})

$null = $window.ShowDialog()
