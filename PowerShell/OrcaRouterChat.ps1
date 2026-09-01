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
$statusText = $window.FindName('StatusText')
$historyStatusText = $window.FindName('HistoryStatusText')
$developerBox = $window.FindName('DeveloperBox')
$newChatButton = $window.FindName('NewChatButton')
$promptExampleBox = $window.FindName('PromptExampleBox')
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

    $traceBox.AppendText(($block -join [Environment]::NewLine))
    $traceBox.ScrollToEnd()
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

function Get-ConversationTranscript {
    param(
        [string]$PendingQuestion = '',
        [string]$PendingAnswer = ''
    )

    $lines = [System.Collections.Generic.List[string]]::new()

    foreach ($turn in $script:conversationHistory) {
        $lines.Add('YOU')
        $lines.Add([string]$turn.User)
        $lines.Add('')
        $lines.Add('ASSISTANT')
        $lines.Add([string]$turn.Assistant)
        $lines.Add('')
    }

    if (-not [string]::IsNullOrWhiteSpace($PendingQuestion)) {
        $lines.Add('YOU')
        $lines.Add($PendingQuestion)
        $lines.Add('')

        if (-not [string]::IsNullOrWhiteSpace($PendingAnswer)) {
            $lines.Add('ASSISTANT')
            $lines.Add($PendingAnswer)
            $lines.Add('')
        }
    }

    if ($lines.Count -eq 0) {
        return 'ここに会話が表示されます。'
    }

    return ($lines -join [Environment]::NewLine).TrimEnd()
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
    Set-ViewModelValue -Name 'Answer' -Value (Get-ConversationTranscript)
}

function Clear-ConversationHistory {
    $script:conversationHistory.Clear()
    $script:currentQuestion = ''
    Update-HistoryStatus
    Set-ViewModelValue -Name 'Answer' -Value (Get-ConversationTranscript)
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

function Set-DeveloperInformation {
    param($WorkerEvent)

    $usage = $WorkerEvent.Usage

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
            $costText = '

    $sendButton.IsEnabled = -not $Busy
    $newChatButton.IsEnabled = -not $Busy
    $apiKeyBox.IsEnabled = -not $Busy
    $modelBox.IsEnabled = -not $Busy
    $modeBox.IsEnabled = -not $Busy
    $promptExampleBox.IsEnabled = -not $Busy

    # The user can prepare the next question while the current request runs.
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

        Set-ViewModelValue -Name 'Answer' -Value "ERROR: $safeMessage"
        Set-ViewModelValue -Name 'StatusText' -Value 'Error - Trace を確認してください'
        $statusText.Foreground = '#B42318'
        $resultTabs.SelectedIndex = 1
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

function Process-OrcaRouterWorkerEvents {
    if ($null -eq $script:workerEventQueue) {
        return
    }

    $workerEvent = $null

    while ($script:workerEventQueue.TryDequeue([ref]$workerEvent)) {
        switch ([string]$workerEvent.Type) {
            'Trace' {
                Add-Trace -Step ([string]$workerEvent.Step) -Direction ([string]$workerEvent.Direction) -Title ([string]$workerEvent.Title) -Data $workerEvent.Data
            }

            'Answer' {
                Set-ViewModelValue -Name 'Answer' -Value ([string]$workerEvent.Text)
                $answerBox.ScrollToEnd()
            }

            'Completed' {
                Set-ViewModelValue -Name 'Answer' -Value ([string]$workerEvent.Answer)
                Set-ViewModelValue -Name 'StatusText' -Value 'Completed'
                $statusText.Foreground = '#0F766E'
                Set-UiBusy -Busy $false
                $script:workerFinishedInUi = $true
            }

            'Error' {
                $safeMessage = Protect-LocalTraceText -Text ([string]$workerEvent.Message)

                Add-Trace -Step 'ERROR' -Direction 'ERROR' -Title 'バックグラウンド処理でエラー' -Data @{
                    Message = $safeMessage
                    Type = [string]$workerEvent.ExceptionType
                }

                Set-ViewModelValue -Name 'Answer' -Value "ERROR: $safeMessage"
                Set-ViewModelValue -Name 'StatusText' -Value 'Error - Trace を確認してください'
                $statusText.Foreground = '#B42318'
                $resultTabs.SelectedIndex = 1
                Set-UiBusy -Busy $false
                $script:workerFinishedInUi = $true
            }
        }

        $workerEvent = $null
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
        [string]$Mode
    )

    Stop-OrcaRouterWorker

    $script:workerEventQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:workerPowerShell = [System.Management.Automation.PowerShell]::Create()

    [void]$script:workerPowerShell.AddCommand($workerPath)
    [void]$script:workerPowerShell.AddParameter('ApiKey', $ApiKey)
    [void]$script:workerPowerShell.AddParameter('Model', $Model)
    [void]$script:workerPowerShell.AddParameter('Question', $Question)
    [void]$script:workerPowerShell.AddParameter('Mode', $Mode)
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
        $safeMessage = Protect-LocalTraceText -Text $_.Exception.Message
        Set-ViewModelValue -Name 'Answer' -Value "ERROR: $safeMessage"
        Set-ViewModelValue -Name 'StatusText' -Value 'Error - 入力値を確認してください'
        $statusText.Foreground = '#B42318'
        return
    }

    $resultTabs.SelectedIndex = 0
    Set-ViewModelValue -Name 'Answer' -Value ''
    Set-ViewModelValue -Name 'StatusText' -Value 'Processing...'
    $statusText.Foreground = '#64748B'
    Set-UiBusy -Busy $true

    Add-Trace -Step 'ASYNC' -Direction 'LOCAL' -Title 'バックグラウンドRunspaceを開始' -Data @{
        Mode = $mode
        UIThread = 'WPF Dispatcher remains responsive'
        ViewModel = 'DataRowView / INotifyPropertyChanged'
    }

    Start-OrcaRouterWorker -ApiKey $apiKey -Model $model -Question $question -Mode $mode
}

$script:workerPollTimer.Add_Tick({
    Process-OrcaRouterWorkerEvents
})

$sendButton.Add_Click({
    Invoke-OrcaRouterChat
})

$clearTraceButton.Add_Click({
    $traceBox.Clear()
    Set-ViewModelValue -Name 'StatusText' -Value 'Ready'
    $statusText.Foreground = '#64748B'
})

$modeBox.Add_SelectionChanged({
    $mode = Get-SelectedMode

    if ($mode -eq 'Tool Calling') {
        $questionBox.Text = 'calculate_sum ツールを使って 123 と 456 を足し、その結果を日本語で説明してください。'
    }
    elseif ($mode -eq 'Streaming') {
        $questionBox.Text = '日本語で「こんにちは。Streamingのテストです。」と短く答えてください。'
    }
    elseif ($mode -eq 'Chat') {
        $questionBox.Text = '日本語で「こんにちは。PowerShell版Chatのテストです。」とだけ答えてください。'
    }
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

        # At a shorter window height, the whole page must remain reachable
        # through the right-side vertical scrollbar.
        $window.Height = 720
        $window.UpdateLayout()

        if ($pageScrollViewer.ScrollableHeight -le 0) {
            throw 'PageScrollViewer must have a scrollable vertical range.'
        }

        $pageScrollViewer.ScrollToEnd()
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ($pageScrollViewer.VerticalOffset -le 0) {
            throw 'PageScrollViewer could not scroll to the lower content.'
        }

        $pageScrollViewer.ScrollToHome()
        $window.Height = 900
        $window.UpdateLayout()

        if ($questionBox.IsReadOnly) { throw 'QuestionBox must be editable.' }
        if (-not $questionBox.IsEnabled) { throw 'QuestionBox must be enabled.' }
        if (-not $questionBox.Focusable) { throw 'QuestionBox must be focusable.' }

        if ($questionBox.ActualHeight -lt 80) {
            throw "QuestionBox is too small: $($questionBox.ActualHeight)"
        }

        if ($null -eq $resultTabs) {
            throw 'ResultTabs was not found.'
        }

        if ($resultTabs.Items.Count -ne 2) {
            throw "ResultTabs must contain Answer and Trace tabs."
        }

        $resultTabs.SelectedIndex = 0
        $window.UpdateLayout()

        if ($answerBox.ActualHeight -lt 180) {
            throw "Answer tab is too small: $($answerBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 1
        $window.UpdateLayout()

        if ($traceBox.ActualHeight -lt 180) {
            throw "Trace tab is too small: $($traceBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 0
        $window.UpdateLayout()

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

$window.Add_ContentRendered({
    [void]$questionBox.Focus()
    $questionBox.CaretIndex = $questionBox.Text.Length
})

$window.Add_Closing({
    Stop-OrcaRouterWorker
})

$null = $window.ShowDialog()
 + ([double]$usage.cost_usd).ToString(
                '0.000000',
                [Globalization.CultureInfo]::InvariantCulture
            )
        }
    }

    $requestText = '{}'
    if ($null -ne $WorkerEvent.Request) {
        $requestText = ConvertTo-TraceText -Data $WorkerEvent.Request
    }

    $responseText = '{}'
    if ($null -ne $WorkerEvent.Response) {
        $responseText = ConvertTo-TraceText -Data $WorkerEvent.Response
    }

    $developerBox.Text = @(
        'Developer Information'
        ('=' * 72)
        "HTTP Status      : $($WorkerEvent.HttpStatus)"
        "Elapsed          : $($WorkerEvent.TotalElapsedMs) ms"
        "Model            : $($WorkerEvent.ActualModel)"
        "Prompt Tokens    : $promptTokens"
        "Completion Tokens: $completionTokens"
        "Total Tokens     : $totalTokens"
        "Cost             : $costText"
        "History          : $($script:conversationHistory.Count) / $($script:MaxHistoryTurns) turns"
        ''
        '--- Request JSON ---'
        $requestText
        ''
        '--- Response JSON ---'
        $responseText
    ) -join [Environment]::NewLine
}

function Set-UiBusy {
    param([bool]$Busy)

    $sendButton.IsEnabled = -not $Busy
    $apiKeyBox.IsEnabled = -not $Busy
    $modelBox.IsEnabled = -not $Busy
    $modeBox.IsEnabled = -not $Busy

    # The user can prepare the next question while the current request runs.
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

        Set-ViewModelValue -Name 'Answer' -Value "ERROR: $safeMessage"
        Set-ViewModelValue -Name 'StatusText' -Value 'Error - Trace を確認してください'
        $statusText.Foreground = '#B42318'
        $resultTabs.SelectedIndex = 1
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

function Process-OrcaRouterWorkerEvents {
    if ($null -eq $script:workerEventQueue) {
        return
    }

    $workerEvent = $null

    while ($script:workerEventQueue.TryDequeue([ref]$workerEvent)) {
        switch ([string]$workerEvent.Type) {
            'Trace' {
                Add-Trace -Step ([string]$workerEvent.Step) -Direction ([string]$workerEvent.Direction) -Title ([string]$workerEvent.Title) -Data $workerEvent.Data
            }

            'Answer' {
                Set-ViewModelValue -Name 'Answer' -Value ([string]$workerEvent.Text)
                $answerBox.ScrollToEnd()
            }

            'Completed' {
                Set-ViewModelValue -Name 'Answer' -Value ([string]$workerEvent.Answer)
                Set-ViewModelValue -Name 'StatusText' -Value 'Completed'
                $statusText.Foreground = '#0F766E'
                Set-UiBusy -Busy $false
                $script:workerFinishedInUi = $true
            }

            'Error' {
                $safeMessage = Protect-LocalTraceText -Text ([string]$workerEvent.Message)

                Add-Trace -Step 'ERROR' -Direction 'ERROR' -Title 'バックグラウンド処理でエラー' -Data @{
                    Message = $safeMessage
                    Type = [string]$workerEvent.ExceptionType
                }

                Set-ViewModelValue -Name 'Answer' -Value "ERROR: $safeMessage"
                Set-ViewModelValue -Name 'StatusText' -Value 'Error - Trace を確認してください'
                $statusText.Foreground = '#B42318'
                $resultTabs.SelectedIndex = 1
                Set-UiBusy -Busy $false
                $script:workerFinishedInUi = $true
            }
        }

        $workerEvent = $null
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
        [string]$Mode
    )

    Stop-OrcaRouterWorker

    $script:workerEventQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:workerPowerShell = [System.Management.Automation.PowerShell]::Create()

    [void]$script:workerPowerShell.AddCommand($workerPath)
    [void]$script:workerPowerShell.AddParameter('ApiKey', $ApiKey)
    [void]$script:workerPowerShell.AddParameter('Model', $Model)
    [void]$script:workerPowerShell.AddParameter('Question', $Question)
    [void]$script:workerPowerShell.AddParameter('Mode', $Mode)
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
        $safeMessage = Protect-LocalTraceText -Text $_.Exception.Message
        Set-ViewModelValue -Name 'Answer' -Value "ERROR: $safeMessage"
        Set-ViewModelValue -Name 'StatusText' -Value 'Error - 入力値を確認してください'
        $statusText.Foreground = '#B42318'
        return
    }

    $resultTabs.SelectedIndex = 0
    Set-ViewModelValue -Name 'Answer' -Value ''
    Set-ViewModelValue -Name 'StatusText' -Value 'Processing...'
    $statusText.Foreground = '#64748B'
    Set-UiBusy -Busy $true

    Add-Trace -Step 'ASYNC' -Direction 'LOCAL' -Title 'バックグラウンドRunspaceを開始' -Data @{
        Mode = $mode
        UIThread = 'WPF Dispatcher remains responsive'
        ViewModel = 'DataRowView / INotifyPropertyChanged'
    }

    Start-OrcaRouterWorker -ApiKey $apiKey -Model $model -Question $question -Mode $mode
}

$script:workerPollTimer.Add_Tick({
    Process-OrcaRouterWorkerEvents
})

$sendButton.Add_Click({
    Invoke-OrcaRouterChat
})

$clearTraceButton.Add_Click({
    $traceBox.Clear()
    Set-ViewModelValue -Name 'StatusText' -Value 'Ready'
    $statusText.Foreground = '#64748B'
})

$modeBox.Add_SelectionChanged({
    $mode = Get-SelectedMode

    if ($mode -eq 'Tool Calling') {
        $questionBox.Text = 'calculate_sum ツールを使って 123 と 456 を足し、その結果を日本語で説明してください。'
    }
    elseif ($mode -eq 'Streaming') {
        $questionBox.Text = '日本語で「こんにちは。Streamingのテストです。」と短く答えてください。'
    }
    elseif ($mode -eq 'Chat') {
        $questionBox.Text = '日本語で「こんにちは。PowerShell版Chatのテストです。」とだけ答えてください。'
    }
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

        # At a shorter window height, the whole page must remain reachable
        # through the right-side vertical scrollbar.
        $window.Height = 720
        $window.UpdateLayout()

        if ($pageScrollViewer.ScrollableHeight -le 0) {
            throw 'PageScrollViewer must have a scrollable vertical range.'
        }

        $pageScrollViewer.ScrollToEnd()
        $window.Dispatcher.Invoke(
            [System.Action]{ },
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        if ($pageScrollViewer.VerticalOffset -le 0) {
            throw 'PageScrollViewer could not scroll to the lower content.'
        }

        $pageScrollViewer.ScrollToHome()
        $window.Height = 900
        $window.UpdateLayout()

        if ($questionBox.IsReadOnly) { throw 'QuestionBox must be editable.' }
        if (-not $questionBox.IsEnabled) { throw 'QuestionBox must be enabled.' }
        if (-not $questionBox.Focusable) { throw 'QuestionBox must be focusable.' }

        if ($questionBox.ActualHeight -lt 80) {
            throw "QuestionBox is too small: $($questionBox.ActualHeight)"
        }

        if ($null -eq $resultTabs) {
            throw 'ResultTabs was not found.'
        }

        if ($resultTabs.Items.Count -ne 2) {
            throw "ResultTabs must contain Answer and Trace tabs."
        }

        $resultTabs.SelectedIndex = 0
        $window.UpdateLayout()

        if ($answerBox.ActualHeight -lt 180) {
            throw "Answer tab is too small: $($answerBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 1
        $window.UpdateLayout()

        if ($traceBox.ActualHeight -lt 180) {
            throw "Trace tab is too small: $($traceBox.ActualHeight)"
        }

        $resultTabs.SelectedIndex = 0
        $window.UpdateLayout()

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

$window.Add_ContentRendered({
    [void]$questionBox.Focus()
    $questionBox.CaretIndex = $questionBox.Text.Length
})

$window.Add_Closing({
    Stop-OrcaRouterWorker
})

$null = $window.ShowDialog()
