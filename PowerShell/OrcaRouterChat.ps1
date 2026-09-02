#requires -Version 5.1
<#
.SYNOPSIS
    OrcaRouter API learning sample using PowerShell + WPF/XAML.

.DESCRIPTION
    The WPF window stays on the UI thread.
    OrcaRouter HTTP work runs in OrcaRouterWorker.ps1 on a background
    PowerShell runspace.

    Model / Mode / Answer / StatusText use WPF Data Binding.
    The ViewModel is a pure .NET DataRowView created from PowerShell.
    DataRowView implements INotifyPropertyChanged, so no C# ViewModel is
    compiled from embedded C# source.

    QuestionBox remains a direct WPF TextBox for reliable keyboard/IME input.
    Its text is mirrored into the ViewModel for state/diagnostics.

    A real API key is never written to the trace.
#>

param(
    [switch]$SyntaxCheck,
    [switch]$UiBindingCheck
)

if ($SyntaxCheck) {
    exit 0
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$script:ApiEndpoint = 'https://api.orcarouter.ai/v1/chat/completions'
$script:ApiKeyPlaceholder = 'xxx-your-orcarouter-api-key-xxx'
$script:DefaultApiKey = 'xxx-your-orcarouter-api-key-xxx'
$script:ReferralUrl = 'https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d'
$script:MaxHistoryTurns = 10

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

function Protect-LocalTraceText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    $sanitized = $Text

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

$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
if (-not (Test-Path -LiteralPath $xamlPath)) {
    throw 'MainWindow.xaml was not found next to OrcaRouterChat.ps1.'
}

$workerPath = Join-Path $PSScriptRoot 'OrcaRouterWorker.ps1'
if (-not (Test-Path -LiteralPath $workerPath)) {
    throw 'OrcaRouterWorker.ps1 was not found next to OrcaRouterChat.ps1.'
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
$resultTabs = $window.FindName('ResultTabs')
$pageScrollViewer = $window.FindName('PageScrollViewer')
$busyProgressBar = $window.FindName('BusyProgressBar')
$statusText = $window.FindName('StatusText')
$historyStatusText = $window.FindName('HistoryStatusText')
$developerBox = $window.FindName('DeveloperBox')
$newChatButton = $window.FindName('NewChatButton')
$promptExampleBox = $window.FindName('PromptExampleBox')
$applyPromptExampleButton = $window.FindName('ApplyPromptExampleButton')
$referralButton = $window.FindName('ReferralButton')
$firstRunPanel = $window.FindName('FirstRunPanel')

# Pure PowerShell/.NET ViewModel.
# DataRowView implements INotifyPropertyChanged and works well with WPF Binding.
Add-Type -AssemblyName System.Data

$viewModelTable = [System.Data.DataTable]::new('OrcaRouterViewModel')
[void]$viewModelTable.Columns.Add('Model', [string])
[void]$viewModelTable.Columns.Add('Mode', [string])
[void]$viewModelTable.Columns.Add('Question', [string])
[void]$viewModelTable.Columns.Add('Answer', [string])
[void]$viewModelTable.Columns.Add('StatusText', [string])

$viewModelRow = $viewModelTable.NewRow()
$viewModelRow['Model'] = 'orcarouter/free'
$viewModelRow['Mode'] = 'Chat'
$viewModelRow['Question'] =
    '日本語で「こんにちは。PowerShell版Chatのテストです。」とだけ答えてください。'
$viewModelRow['Answer'] = 'ここに回答が表示されます。'
$viewModelRow['StatusText'] = 'Ready'
$viewModelTable.Rows.Add($viewModelRow)

$viewModel = $viewModelTable.DefaultView[0]

$window.DataContext = $viewModel

$modelBox.DataContext = $viewModel
$modeBox.DataContext = $viewModel
$answerBox.DataContext = $viewModel
$statusText.DataContext = $viewModel

# Question is intentionally the actual TextBox input source.
$questionBox.Text = [string]$viewModel.Row['Question']

$apiKeyBox.Password = $script:DefaultApiKey

$script:workerPowerShell = $null
$script:workerAsyncResult = $null
$script:workerEventQueue = $null
$script:workerFinishedInUi = $false
$script:currentQuestion = ''
$script:conversationHistory = [System.Collections.ArrayList]::new()
$script:traceBuffer = [System.Text.StringBuilder]::new()
$script:latestDeveloperEvent = $null

$script:workerPollTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:workerPollTimer.Interval = [TimeSpan]::FromMilliseconds(50)

function Set-ViewModelValue {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value
    )

    $viewModel.Row[$Name] = $Value
}

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
    $blockText = ($block -join [Environment]::NewLine) + [Environment]::NewLine
    [void]$script:traceBuffer.Append($blockText)

    # Keep diagnostics separate from the main answer view. Updating a large,
    # hidden TextBox for every trace event can make long requests feel frozen.
    if ($resultTabs.SelectedIndex -eq 2) {
        $traceBox.AppendText($blockText)
        $traceBox.ScrollToEnd()
    }
}

function Test-ConfiguredApiKey {
    $value = $apiKeyBox.Password.Trim()

    return (
        -not [string]::IsNullOrWhiteSpace($value) -and
        $value -ne $script:ApiKeyPlaceholder -and
        -not $value.StartsWith('xxx-')
    )
}

function Update-FirstRunPanel {
    if (Test-ConfiguredApiKey) {
        $firstRunPanel.Visibility = [System.Windows.Visibility]::Collapsed
    }
    else {
        $firstRunPanel.Visibility = [System.Windows.Visibility]::Visible
    }
}

function Update-HistoryStatus {
    $count = $script:conversationHistory.Count
    $historyStatusText.Text = "履歴 $count / $($script:MaxHistoryTurns) 往復"
}

function Add-ConversationTurn {
    param(
        [Parameter(Mandatory = $true)][string]$Question,
        [Parameter(Mandatory = $true)][string]$Assistant
    )

    [void]$script:conversationHistory.Add(
        [pscustomobject]@{
            User = $Question
            Assistant = $Assistant
        }
    )

    while ($script:conversationHistory.Count -gt $script:MaxHistoryTurns) {
        $script:conversationHistory.RemoveAt(0)
    }

    Update-HistoryStatus

    # The primary Answer area shows only the latest assistant response.
    # Conversation history is kept internally only for the next API request.
    Set-ViewModelValue -Name 'Answer' -Value $Assistant
}

function Clear-ConversationHistory {
    $script:conversationHistory.Clear()
    $script:currentQuestion = ''
    Update-HistoryStatus
    Set-ViewModelValue -Name 'Answer' -Value 'ここに回答が表示されます。'
    Set-ViewModelValue -Name 'StatusText' -Value 'New chat - 履歴をクリアしました'
    $statusText.Foreground = '#64748B'
}

function Get-HistorySnapshot {
    $snapshot = @()

    foreach ($turn in $script:conversationHistory) {
        $snapshot += [pscustomobject]@{
            User = [string]$turn.User
            Assistant = [string]$turn.Assistant
        }
    }

    return $snapshot
}

function Get-WorkerEventValue {
    param(
        $WorkerEvent,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue = $null
    )

    if ($null -eq $WorkerEvent) {
        return $DefaultValue
    }

    $property = $WorkerEvent.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function Set-DeveloperInformation {
    param($WorkerEvent)

    $usage = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'Usage'
    $httpStatus = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'HttpStatus' -DefaultValue '-'
    $elapsedMs = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'TotalElapsedMs' -DefaultValue '-'
    $actualModel = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'ActualModel' -DefaultValue ([string]$viewModel.Row['Model'])
    $requestValue = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'Request'
    $responseValue = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'Response'
    $errorMessage = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'Message' -DefaultValue ''
    $exceptionType = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'ExceptionType' -DefaultValue ''
    $scriptLineNumber = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'ScriptLineNumber' -DefaultValue ''
    $positionMessage = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'PositionMessage' -DefaultValue ''
    $scriptStackTrace = Get-WorkerEventValue -WorkerEvent $WorkerEvent -Name 'ScriptStackTrace' -DefaultValue ''

    $promptTokens = '-'
    $completionTokens = '-'
    $totalTokens = '-'
    $costText = '(not returned)'

    if ($null -ne $usage) {
        if ($null -ne $usage.PSObject.Properties['prompt_tokens']) {
            $promptTokens = [string]$usage.prompt_tokens
        }
        if ($null -ne $usage.PSObject.Properties['completion_tokens']) {
            $completionTokens = [string]$usage.completion_tokens
        }
        if ($null -ne $usage.PSObject.Properties['total_tokens']) {
            $totalTokens = [string]$usage.total_tokens
        }
        if ($null -ne $usage.PSObject.Properties['cost_usd']) {
            $costText = [string][char]36 + ([double]$usage.cost_usd).ToString(
                '0.000000',
                [Globalization.CultureInfo]::InvariantCulture
            )
        }
    }

    $requestText = '{}'
    if ($null -ne $requestValue) {
        $requestText = ConvertTo-TraceText -Data $requestValue
    }

    $responseText = '{}'
    if ($null -ne $responseValue) {
        $responseText = ConvertTo-TraceText -Data $responseValue
    }

    $lines = @(
        'Developer Information'
        ('=' * 72)
        "HTTP Status      : $httpStatus"
        "Elapsed          : $elapsedMs ms"
        "Model            : $actualModel"
        "Prompt Tokens    : $promptTokens"
        "Completion Tokens: $completionTokens"
        "Total Tokens     : $totalTokens"
        "Cost             : $costText"
        "History          : $($script:conversationHistory.Count) / $($script:MaxHistoryTurns) turns"
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$errorMessage)) {
        $lines += ''
        $lines += '--- Error ---'
        $lines += (Protect-LocalTraceText -Text ([string]$errorMessage))

        if (-not [string]::IsNullOrWhiteSpace([string]$exceptionType)) {
            $lines += "Type: $exceptionType"
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$scriptLineNumber)) {
            $lines += "Script line: $scriptLineNumber"
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$positionMessage)) {
            $lines += (Protect-LocalTraceText -Text ([string]$positionMessage))
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$scriptStackTrace)) {
            $lines += '--- Script stack ---'
            $lines += (Protect-LocalTraceText -Text ([string]$scriptStackTrace))
        }
    }

    $lines += ''
    $lines += '--- Request JSON ---'
    $lines += $requestText
    $lines += ''
    $lines += '--- Response JSON / Error body ---'
    $lines += $responseText

    $developerBox.Text = $lines -join [Environment]::NewLine
}

function Set-DeveloperPendingInformation {
    param([string]$Model)

    # A new request must not reopen the previous request's diagnostics.
    $script:latestDeveloperEvent = $null

    $developerBox.Text = @(
        'Developer Information'
        ('=' * 72)
        'Result           : SENDING'
        'HTTP Status      : -'
        'Elapsed          : -'
        "Model            : $Model"
        'Prompt Tokens    : -'
        'Completion Tokens: -'
        'Total Tokens     : -'
        'Cost             : (not returned)'
        "History          : $($script:conversationHistory.Count) / $($script:MaxHistoryTurns) turns"
        ''
        'Request / Response will appear when the worker reports them.'
    ) -join [Environment]::NewLine
}
function Store-DeveloperInformation {
    param($WorkerEvent)

    $script:latestDeveloperEvent = $WorkerEvent
    $developerBox.Text = @(
        'Developer Information'
        ('=' * 72)
        '診断情報を取得しました。Developer タブを開くと Request / Response を展開します。'
    ) -join [Environment]::NewLine

    if ($resultTabs.SelectedIndex -eq 1) {
        Set-DeveloperInformation -WorkerEvent $WorkerEvent
    }
}

function Show-RequestError {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$ExceptionType = '',
        $WorkerEvent = $null
    )

    $safeMessage = Protect-LocalTraceText -Text $Message
    Set-ViewModelValue -Name 'Answer' -Value ("ERROR: " + $safeMessage)

    if ($null -eq $WorkerEvent) {
        $WorkerEvent = [pscustomobject]@{
            Message = $safeMessage
            ExceptionType = $ExceptionType
            ActualModel = [string]$viewModel.Row['Model']
        }
    }

    Store-DeveloperInformation -WorkerEvent $WorkerEvent
    Set-ViewModelValue -Name 'StatusText' -Value 'Error - エラー内容を回答欄に表示しました'
    $statusText.Foreground = '#B42318'

    # The primary result remains visible; Developer / Trace are optional diagnostics.
    $resultTabs.SelectedIndex = 0
}
function Set-UiBusy {
    param([bool]$Busy)

    $sendButton.IsEnabled = -not $Busy
    $sendButton.Content = if ($Busy) { '回答待ち...' } else { '送信' }
    $newChatButton.IsEnabled = -not $Busy
    $apiKeyBox.IsEnabled = -not $Busy
    $modelBox.IsEnabled = -not $Busy
    $modeBox.IsEnabled = -not $Busy
    $promptExampleBox.IsEnabled = -not $Busy
    $applyPromptExampleButton.IsEnabled = -not $Busy
    $busyProgressBar.Visibility = if ($Busy) {
        [System.Windows.Visibility]::Visible
    }
    else {
        [System.Windows.Visibility]::Collapsed
    }

    # The current HTTP request runs in a background runspace. Keep the editor
    # and the window itself responsive while the user waits for the answer.
    $questionBox.IsEnabled = $true
}

function Get-SelectedMode {
    $mode = [string]$viewModel.Row['Mode']

    if ([string]::IsNullOrWhiteSpace($mode)) {
        return 'Chat'
    }

    return $mode
}

function Assert-Inputs {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [string]$Mode
    )

    Add-Trace -Step 'STEP 1' -Direction 'LOCAL' -Title '入力値を検証' -Data @{
        ApiKeyMasked = Mask-ApiKey -ApiKey $ApiKey
        Model = $Model
        Mode = $Mode
        QuestionChars = $Question.Length
    }

    if (
        [string]::IsNullOrWhiteSpace($ApiKey) -or
        $ApiKey -eq $script:ApiKeyPlaceholder -or
        $ApiKey.StartsWith('xxx-')
    ) {
        throw 'APIキーがダミー値のままです。OrcaRouterで発行したAPIキーを入力してください。'
    }

    if ([string]::IsNullOrWhiteSpace($Model)) {
        throw 'Model を入力してください。'
    }

    if ([string]::IsNullOrWhiteSpace($Question)) {
        throw '質問を入力してください。'
    }
}

function Stop-OrcaRouterWorker {
    if ($script:workerPollTimer.IsEnabled) {
        $script:workerPollTimer.Stop()
    }

    if ($null -ne $script:workerPowerShell) {
        try {
            if (
                $null -ne $script:workerAsyncResult -and
                -not $script:workerAsyncResult.IsCompleted
            ) {
                $script:workerPowerShell.Stop()
            }
        }
        catch {
            # Best-effort cleanup.
        }

        try {
            $script:workerPowerShell.Dispose()
        }
        catch {
            # Ignore cleanup-only errors.
        }
    }

    $script:workerPowerShell = $null
    $script:workerAsyncResult = $null
    $script:workerEventQueue = $null
    $script:workerFinishedInUi = $false
}

function Complete-OrcaRouterWorker {
    if ($null -eq $script:workerPowerShell) {
        return
    }

    try {
        if ($null -ne $script:workerAsyncResult) {
            [void]$script:workerPowerShell.EndInvoke($script:workerAsyncResult)
        }
    }
    catch {
        $safeMessage = Protect-LocalTraceText -Text $_.Exception.Message

        Add-Trace -Step 'ERROR' -Direction 'ERROR' -Title 'バックグラウンドRunspaceの終了処理でエラー' -Data @{
            Message = $safeMessage
        }

        Show-RequestError -Message $safeMessage -ExceptionType $_.Exception.GetType().FullName
    }
    finally {
        try {
            $script:workerPowerShell.Dispose()
        }
        catch {
            # Ignore cleanup-only errors.
        }

        $script:workerPowerShell = $null
        $script:workerAsyncResult = $null
        $script:workerEventQueue = $null

        if (-not $script:workerFinishedInUi) {
            Set-UiBusy -Busy $false
        }

        $script:workerFinishedInUi = $false
        $script:workerPollTimer.Stop()
    }
}

function Show-StreamingAnswer {
    param([string]$Text)

    Set-ViewModelValue -Name 'Answer' -Value $Text
    Set-ViewModelValue -Name 'StatusText' -Value '回答受信中...'
    $statusText.Foreground = '#64748B'
    $answerBox.ScrollToEnd()
}

function Process-OrcaRouterWorkerEvents {
    if ($null -eq $script:workerEventQueue) {
        return
    }

    # Do not drain an arbitrarily large queue in one Dispatcher tick.
    # Long Streaming responses can otherwise monopolize the WPF UI thread.
    $maxEventsPerTick = 40
    $processed = 0
    $workerEvent = $null
    $latestAnswer = $null

    while (
        $processed -lt $maxEventsPerTick -and
        $script:workerEventQueue.TryDequeue([ref]$workerEvent)
    ) {
        $processed += 1

        switch ([string]$workerEvent.Type) {
            'Trace' {
                Add-Trace -Step ([string]$workerEvent.Step) -Direction ([string]$workerEvent.Direction) -Title ([string]$workerEvent.Title) -Data $workerEvent.Data
            }

            'Answer' {
                # Coalesce multiple Streaming updates and paint only the latest
                # value once per UI tick.
                $latestAnswer = [string]$workerEvent.Text
            }

            'Completed' {
                if ($null -ne $latestAnswer) {
                    Show-StreamingAnswer -Text $latestAnswer
                    $latestAnswer = $null
                }

                Add-ConversationTurn -Question $script:currentQuestion -Assistant ([string]$workerEvent.Answer)
                Store-DeveloperInformation -WorkerEvent $workerEvent
                Set-ViewModelValue -Name 'StatusText' -Value 'Completed'
                $statusText.Foreground = '#0F766E'

                $resultTabs.SelectedIndex = 0
                Set-UiBusy -Busy $false
                $script:workerFinishedInUi = $true
            }

            'Error' {
                $latestAnswer = $null
                $safeMessage = Protect-LocalTraceText -Text ([string](Get-WorkerEventValue -WorkerEvent $workerEvent -Name 'Message' -DefaultValue 'Unknown error'))
                $exceptionType = [string](Get-WorkerEventValue -WorkerEvent $workerEvent -Name 'ExceptionType' -DefaultValue '(unknown)')

                Add-Trace -Step 'ERROR' -Direction 'ERROR' -Title 'バックグラウンド処理でエラー' -Data @{
                    Message = $safeMessage
                    Type = $exceptionType
                    HttpStatus = Get-WorkerEventValue -WorkerEvent $workerEvent -Name 'HttpStatus' -DefaultValue '-'
                    ScriptLineNumber = Get-WorkerEventValue -WorkerEvent $workerEvent -Name 'ScriptLineNumber' -DefaultValue '-'
                    PositionMessage = Get-WorkerEventValue -WorkerEvent $workerEvent -Name 'PositionMessage' -DefaultValue ''
                    ScriptStackTrace = Get-WorkerEventValue -WorkerEvent $workerEvent -Name 'ScriptStackTrace' -DefaultValue ''
                }

                Show-RequestError -Message $safeMessage -ExceptionType $exceptionType -WorkerEvent $workerEvent
                Set-UiBusy -Busy $false
                $script:workerFinishedInUi = $true
            }
        }

        $workerEvent = $null
    }

    if ($null -ne $latestAnswer) {
        Show-StreamingAnswer -Text $latestAnswer
    }

    if (
        $null -ne $script:workerAsyncResult -and
        $script:workerAsyncResult.IsCompleted -and
        $script:workerEventQueue.IsEmpty
    ) {
        Complete-OrcaRouterWorker
    }
}

function Start-OrcaRouterWorker {
    param(
        [string]$ApiKey,
        [string]$Model,
        [string]$Question,
        [string]$Mode,
        [object[]]$History
    )

    Stop-OrcaRouterWorker

    $script:workerEventQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:workerPowerShell = [System.Management.Automation.PowerShell]::Create()

    [void]$script:workerPowerShell.AddCommand($workerPath)
    [void]$script:workerPowerShell.AddParameter('ApiKey', $ApiKey)
    [void]$script:workerPowerShell.AddParameter('Model', $Model)
    [void]$script:workerPowerShell.AddParameter('Question', $Question)
    [void]$script:workerPowerShell.AddParameter('Mode', $Mode)
    [void]$script:workerPowerShell.AddParameter('History', $History)
    [void]$script:workerPowerShell.AddParameter('EventQueue', $script:workerEventQueue)

    $script:workerFinishedInUi = $false
    $script:workerAsyncResult = $script:workerPowerShell.BeginInvoke()
    $script:workerPollTimer.Start()
}

function Invoke-OrcaRouterChat {
    if (
        $null -ne $script:workerAsyncResult -and
        -not $script:workerAsyncResult.IsCompleted
    ) {
        return
    }

    $apiKey = $apiKeyBox.Password.Trim()
    $model = [string]$viewModel.Row['Model']
    $model = $model.Trim()
    $question = $questionBox.Text.Trim()
    $mode = Get-SelectedMode

    try {
        Assert-Inputs -ApiKey $apiKey -Model $model -Question $question -Mode $mode
    }
    catch {
        $script:currentQuestion = $question
        $safeMessage = Protect-LocalTraceText -Text $_.Exception.Message
        Show-RequestError -Message $safeMessage -ExceptionType $_.Exception.GetType().FullName
        return
    }

    $resultTabs.SelectedIndex = 0
    $script:currentQuestion = $question

    # The main result area is answer-only. Clear the previous answer and use
    # the always-visible status row to show that the request is in progress.
    Set-ViewModelValue -Name 'Answer' -Value ''
    Set-ViewModelValue -Name 'StatusText' -Value '送信済み・回答待ち...'
    $statusText.Foreground = '#64748B'
    Set-DeveloperPendingInformation -Model $model
    Set-UiBusy -Busy $true

    Add-Trace -Step 'ASYNC' -Direction 'LOCAL' -Title 'バックグラウンドRunspaceを開始' -Data @{
        Mode = $mode
        UIThread = 'WPF Dispatcher remains responsive'
        ViewModel = 'DataRowView / INotifyPropertyChanged'
    }

    Start-OrcaRouterWorker -ApiKey $apiKey -Model $model -Question $question -Mode $mode -History (Get-HistorySnapshot)
}

$script:workerPollTimer.Add_Tick({
    Process-OrcaRouterWorkerEvents
})

$sendButton.Add_Click({
    Invoke-OrcaRouterChat
})

$clearTraceButton.Add_Click({
    [void]$script:traceBuffer.Clear()
    $traceBox.Clear()
    Set-ViewModelValue -Name 'StatusText' -Value 'Ready'
    $statusText.Foreground = '#64748B'
})

$newChatButton.Add_Click({
    Clear-ConversationHistory
})

$referralButton.Add_Click({
    Start-Process $script:ReferralUrl
})

$apiKeyBox.Add_PasswordChanged({
    Update-FirstRunPanel
})

$applyPromptExampleButton.Add_Click({
    if ($promptExampleBox.SelectedIndex -le 0) {
        Set-ViewModelValue -Name 'StatusText' -Value 'プロンプト例を選択してください'
        $statusText.Foreground = '#64748B'
        return
    }

    $selectedText = [string](($promptExampleBox.SelectedItem).Content)
    $blankLine = [Environment]::NewLine + [Environment]::NewLine

    switch ($selectedText) {
        '要約' {
            $questionBox.Text = '次の文章を3行で要約してください。' + $blankLine + 'ここに文章を貼り付けてください。'
        }
        '初心者向け説明' {
            $questionBox.Text = '次の内容を、専門用語を補足しながら初心者向けに説明してください。' + $blankLine + 'ここに内容を貼り付けてください。'
        }
        'コードレビュー' {
            $questionBox.Text = '次のコードをレビューし、問題点・理由・改善例の順に説明してください。' + $blankLine + 'ここにコードを貼り付けてください。'
        }
        'JSON形式で整理' {
            $questionBox.Text = '次の内容を整理し、JSON形式だけで返してください。' + $blankLine + 'ここに内容を貼り付けてください。'
        }
        '英訳' {
            $questionBox.Text = '次の日本語を自然な英語に翻訳してください。' + $blankLine + 'ここに文章を貼り付けてください。'
        }
    }

    Set-ViewModelValue -Name 'StatusText' -Value ("プロンプト例「" + $selectedText + "」を質問欄に挿入しました")
    $statusText.Foreground = '#64748B'
    [void]$questionBox.Focus()
    $questionBox.CaretIndex = $questionBox.Text.Length
})

$resultTabs.Add_SelectionChanged({
    param($sender, $eventArgs)

    if ($sender -ne $resultTabs) {
        return
    }

    if ($resultTabs.SelectedIndex -eq 1 -and $null -ne $script:latestDeveloperEvent) {
        Set-DeveloperInformation -WorkerEvent $script:latestDeveloperEvent
    }
    elseif ($resultTabs.SelectedIndex -eq 2) {
        $traceBox.Text = $script:traceBuffer.ToString()
        $traceBox.ScrollToEnd()
    }
})

$modeBox.Add_SelectionChanged({
    $mode = Get-SelectedMode

    # Changing Mode must never destroy a question the user is editing.
    Set-ViewModelValue -Name 'StatusText' -Value "Mode: $mode"
    $statusText.Foreground = '#64748B'
})

$questionBox.Add_TextChanged({
    Set-ViewModelValue -Name 'Question' -Value $questionBox.Text
})

$questionBox.Add_PreviewKeyDown({
    param($sender, $eventArgs)

    if (
        $eventArgs.Key -eq [System.Windows.Input.Key]::Enter -and
        (
            [System.Windows.Input.Keyboard]::Modifiers -band
            [System.Windows.Input.ModifierKeys]::Control
        )
    ) {
        $eventArgs.Handled = $true
        Invoke-OrcaRouterChat
    }
})

if ($UiBindingCheck) {
    try {
        $window.WindowState = [System.Windows.WindowState]::Normal
        $window.Width = 1200
        $window.Height = 900
        $window.Show()
        $window.UpdateLayout()

        if ($window.ResizeMode -ne [System.Windows.ResizeMode]::CanResizeWithGrip) {
            throw 'Window must support resize/minimize/maximize.'
        }

        if ($null -eq $pageScrollViewer) {
            throw 'PageScrollViewer was not found.'
        }

        if (
            $pageScrollViewer.VerticalScrollBarVisibility -ne
            [System.Windows.Controls.ScrollBarVisibility]::Auto
        ) {
            throw 'PageScrollViewer must use an automatic vertical scrollbar.'
        }

        if ($null -eq $busyProgressBar) {
            throw 'BusyProgressBar was not found.'
        }

        # At a shorter window height, the entire form must remain reachable
        # through the right-side page scrollbar.
        $window.Height = 720
        $window.UpdateLayout()

        if ($pageScrollViewer.ScrollableHeight -le 0) {
            throw 'PageScrollViewer must provide a vertical scroll range.'
        }

        $pageScrollViewer.ScrollToEnd()
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ($pageScrollViewer.VerticalOffset -le 0) {
            throw 'PageScrollViewer could not scroll to the lower RESULT area.'
        }

        $pageScrollViewer.ScrollToHome()
        $window.Height = 900
        $window.UpdateLayout()

        # Long text must stay inside the Question/Answer editors. The outer
        # window remains resizable instead of growing a page-sized document.
        if (
            $questionBox.VerticalScrollBarVisibility -ne
            [System.Windows.Controls.ScrollBarVisibility]::Auto
        ) {
            throw 'QuestionBox must use its own automatic vertical scrollbar.'
        }

        if (
            $answerBox.VerticalScrollBarVisibility -ne
            [System.Windows.Controls.ScrollBarVisibility]::Auto
        ) {
            throw 'AnswerBox must use its own automatic vertical scrollbar.'
        }


        if ($questionBox.ActualHeight -lt 120) {
            throw "QuestionBox must remain visibly multi-line: $($questionBox.ActualHeight)"
        }

        if ($questionBox.IsReadOnly) { throw 'QuestionBox must be editable.' }
        if (-not $questionBox.IsEnabled) { throw 'QuestionBox must be enabled.' }
        if (-not $questionBox.Focusable) { throw 'QuestionBox must be focusable.' }

        if ($null -eq $resultTabs) {
            throw 'ResultTabs was not found.'
        }

        if ($resultTabs.Items.Count -ne 3) {
            throw "ResultTabs must contain Answer, Developer, and Trace tabs."
        }

        if ($null -eq $developerBox) {
            throw 'DeveloperBox was not found.'
        }

        if ($null -eq $newChatButton) {
            throw 'NewChatButton was not found.'
        }

        if ($null -eq $promptExampleBox) {
            throw 'PromptExampleBox was not found.'
        }

        if ($null -eq $applyPromptExampleButton) {
            throw 'ApplyPromptExampleButton was not found.'
        }

        # Changing Mode must not overwrite the user's question.
        $questionBox.Text = 'Mode change must not overwrite this text'
        $modeBox.SelectedIndex = 1
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ($questionBox.Text -ne 'Mode change must not overwrite this text') {
            throw 'Changing Mode must not overwrite QuestionBox.'
        }

        $modeBox.SelectedIndex = 0

        # Selecting a prompt example must not overwrite the user's question.
        $questionBox.Text = 'Prompt selection must not overwrite this text'
        $promptExampleBox.SelectedIndex = 1
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ($questionBox.Text -ne 'Prompt selection must not overwrite this text') {
            throw 'Selecting PromptExampleBox must not change QuestionBox until Apply is clicked.'
        }

        $resultTabs.SelectedIndex = 0
        $window.UpdateLayout()

        if ($answerBox.ActualHeight -lt 140) {
            throw "Answer tab is too small: $($answerBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 1
        $window.UpdateLayout()

        if ($developerBox.ActualHeight -lt 180) {
            throw "Developer tab is too small: $($developerBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 2
        $window.UpdateLayout()

        if ($traceBox.ActualHeight -lt 180) {
            throw "Trace tab is too small: $($traceBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 0
        $window.UpdateLayout()

        $questionHeightBefore = $questionBox.ActualHeight
        $answerHeightBefore = $answerBox.ActualHeight
        $longText = ('長文テスト ' * 4000)

        $questionBox.Text = $longText
        Set-ViewModelValue -Name 'Answer' -Value $longText
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::DataBind
        )
        $window.UpdateLayout()

        if ([Math]::Abs($questionBox.ActualHeight - $questionHeightBefore) -gt 2) {
            throw 'Long Question text must not expand the whole form.'
        }

        if ([Math]::Abs($answerBox.ActualHeight - $answerHeightBefore) -gt 2) {
            throw 'Long Answer text must not expand the whole form.'
        }

        Set-UiBusy -Busy $true

        if ($busyProgressBar.Visibility -ne [System.Windows.Visibility]::Visible) {
            throw 'Busy indicator must be visible while a request is running.'
        }

        if (-not $questionBox.IsEnabled) {
            throw 'QuestionBox must stay enabled while waiting for an answer.'
        }

        if ($sendButton.IsEnabled) {
            throw 'SendButton must be disabled while a request is running.'
        }

        Set-UiBusy -Busy $false

        if ($busyProgressBar.Visibility -ne [System.Windows.Visibility]::Collapsed) {
            throw 'Busy indicator must collapse after the request finishes.'
        }

        if ($viewModel -isnot [System.ComponentModel.INotifyPropertyChanged]) {
            throw 'ViewModel must implement INotifyPropertyChanged.'
        }

        Set-ViewModelValue -Name 'Answer' -Value 'Binding answer test'

        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::DataBind
        )

        if ($answerBox.Text -ne 'Binding answer test') {
            throw 'Answer Data Binding did not update the TextBox.'
        }

        Set-ViewModelValue -Name 'StatusText' -Value 'Binding status test'

        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::DataBind
        )

        if ($statusText.Text -ne 'Binding status test') {
            throw 'Status Data Binding did not update the TextBlock.'
        }

        $questionBox.Text = 'Question editor input test'

        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ([string]$viewModel.Row['Question'] -ne 'Question editor input test') {
            throw 'QuestionBox TextChanged did not update the ViewModel.'
        }

        # Regression test: multiple successful turns are retained for API
        # context, but the primary Answer area shows only the latest answer.
        $script:conversationHistory.Clear()
        Add-ConversationTurn -Question 'First question' -Assistant 'First answer'
        Add-ConversationTurn -Question 'Second question' -Assistant 'Second answer'
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::DataBind
        )

        if ($answerBox.Text -ne 'Second answer') {
            throw 'Answer must show only the latest assistant response.'
        }

        if ($answerBox.Text -match 'First question|Second question|First answer') {
            throw 'Answer must not concatenate prior conversation turns.'
        }

        if ($script:conversationHistory.Count -ne 2) {
            throw 'Answer-only rendering must not discard internal conversation history.'
        }

        Clear-ConversationHistory

        # Regression test: one Dispatcher tick must not drain an unbounded
        # Streaming event burst.
        $script:workerEventQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        foreach ($index in 1..100) {
            $script:workerEventQueue.Enqueue(
                [pscustomobject]@{
                    Type = 'Answer'
                    Text = "partial $index"
                }
            )
        }

        Process-OrcaRouterWorkerEvents
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::DataBind
        )

        if ($script:workerEventQueue.IsEmpty) {
            throw 'One UI tick drained the entire event burst; responsiveness guard failed.'
        }

        if ($answerBox.Text -ne 'partial 40') {
            throw 'Streaming Answer events were not coalesced to the latest value in the UI tick.'
        }

        $script:workerEventQueue = $null

        # Regression test: an API/worker error must remain visible in Answer,
        # must populate Developer, and must not force the Trace tab.
        $script:currentQuestion = 'Synthetic question'
        $script:workerEventQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        $script:workerAsyncResult = $null
        $script:workerFinishedInUi = $false

        $script:workerEventQueue.Enqueue(
            [pscustomobject]@{
                Type = 'Error'
                Message = 'Synthetic API error'
                ExceptionType = 'System.Exception'
                TotalElapsedMs = 42
                HttpStatus = 429
                Usage = $null
                Request = [pscustomobject]@{ model = 'orcarouter/free' }
                Response = [pscustomobject]@{ error = 'quota' }
                ActualModel = 'orcarouter/free'
            }
        )

        Process-OrcaRouterWorkerEvents
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::DataBind
        )

        if ($answerBox.Text -match 'Synthetic question') {
            throw 'Answer must not repeat the submitted question.'
        }

        if ($answerBox.Text -notmatch 'Synthetic API error') {
            throw 'Error rendering did not show the API error in Answer.'
        }

        if ($resultTabs.SelectedIndex -ne 0) {
            throw 'Error handling must keep the Answer tab selected.'
        }

        # Diagnostics are intentionally separate and rendered only when the
        # Developer tab is opened.
        $resultTabs.SelectedIndex = 1
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ($developerBox.Text -notmatch 'HTTP Status\s+: 429') {
            throw 'Developer Information did not receive error HTTP status.'
        }

        if ($developerBox.Text -notmatch 'Synthetic API error') {
            throw 'Developer Information did not show error details.'
        }

        $resultTabs.SelectedIndex = 0

        $script:workerEventQueue = $null
    }
    finally {
        $window.Close()
    }

    exit 0
}

Add-Trace -Step 'READY' -Direction 'LOCAL' -Title 'サンプルを起動' -Data @{
    Endpoint = $script:ApiEndpoint
    Model = [string]$viewModel.Row['Model']
    Mode = Get-SelectedMode
    ViewModel = 'DataRowView / INotifyPropertyChanged'
    Async = 'Background PowerShell runspace + DispatcherTimer'
    Note = 'APIキーはダミー値です。実行前に画面上で差し替えてください。'
}

Update-FirstRunPanel
Update-HistoryStatus
$developerBox.Text = 'Developer Information will appear after a request.'

$window.Add_ContentRendered({
    [void]$questionBox.Focus()
    $questionBox.CaretIndex = $questionBox.Text.Length
})

$window.Add_Closing({
    Stop-OrcaRouterWorker
})

$null = $window.ShowDialog()
