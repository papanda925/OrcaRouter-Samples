# OrcaRouter PowerShell Sample

PowerShell + WPF/XAML で OrcaRouter の Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** を切り替えて比較できる学習者向けサンプルです。

## Architecture

- `MainWindow.xaml` - View（画面定義）
- `OrcaRouterChat.ps1` - UI、Data Binding、ViewModel、Runspace制御
- `OrcaRouterWorker.ps1` - Chat / Streaming / Tool Calling のHTTP処理

PowerShell版は **C#コードを埋め込まず、PowerShell + .NETだけ** で構成しています。

ViewModelには1行だけの `DataTable` から作った `System.Data.DataRowView` を使います。`DataRowView` は `INotifyPropertyChanged` を実装しているため、PowerShell側でAnswerやStatusなどの列値を変更するとWPFのData Bindingへ変更通知できます。C#のViewModelクラスは使いません。

```text
WPF UI thread
  ├─ MainWindow.xaml
  ├─ Question入力
  ├─ Answer / Status のData Binding
  ├─ 最大化 / リサイズ / GridSplitter
  └─ DispatcherTimer
          ↑
          │ ConcurrentQueue
          ↓
Background PowerShell Runspace
  └─ OrcaRouterWorker.ps1
       ├─ Chat
       ├─ Streaming
       └─ Tool Calling
```

役割は次のように分けています。

| 項目 | 実装 |
|---|---|
| Model / Mode | TwoWay Data Binding |
| Answer / Status | OneWay Data Binding |
| ViewModel変更通知 | `DataRowView` の `INotifyPropertyChanged` |
| Question | WPF `TextBox` を直接入力し、`TextChanged` でViewModelへ同期 |
| API Key | `PasswordBox.Password` を直接取得 |
| Trace | `AppendText()` で追記 |
| HTTP通信 | Background PowerShell Runspace |
| UIへの受け渡し | `ConcurrentQueue` + `DispatcherTimer` |

これは **分かりやすさを優先した軽量MVVM** です。`ICommand` やRelayCommandは使わず、ボタンは `Add_Click({...})` で処理します。

## Data Binding と非同期処理

`INotifyPropertyChanged` と非同期処理は役割が異なります。

```text
Background Runspace
      ↓
OrcaRouterから結果を受信
      ↓
ConcurrentQueue
      ↓
UI側 DispatcherTimer
      ↓
ViewModel.Answer を変更
      ↓
INotifyPropertyChanged
      ↓
Data Binding
      ↓
AnswerBox を更新
```

`INotifyPropertyChanged` は「値が変わったことをWPFへ知らせる仕組み」です。画面を固めない役割は **Background PowerShell Runspace** が担当します。

通常ChatやTool CallingでAPI応答を待っている間も、WPFのUIスレッドではHTTP待機をしません。そのため、ウィンドウの移動・最大化・リサイズ・Question入力などを継続できます。

StreamingではWorker側がSSEを読み取り、途中経過のAnswerイベントをQueueへ渡します。UI側はQueueから受け取った文字列をViewModelの `Answer` へ設定し、Data Bindingで回答欄へ反映します。

## Responsive window layout / page scroll / Result tabs

画面は **ウィンドウサイズに連動して伸縮** しつつ、高さが足りない場合は右側のページ全体用スクロールバーで下まで移動できる構成です。

```text
Window
  └─ PageScrollViewer
       ├─ Header / API settings
       ├─ INPUT
       ├─ GridSplitter
       ├─ RESULT
       │    ├─ 回答 Tab
       │    ├─ Developer Tab
       │    └─ トレース Tab
       └─ Status
```

`PageScrollViewer` は `VerticalScrollBarVisibility="Auto"` です。最大化・元に戻す・画面解像度・Windowsの表示倍率などによって縦方向に収まらない場合だけ、**ウィンドウ右側に縦スクロールバー**が表示されます。これで下部のRESULTやStatusまで移動できます。

RESULTは `TabControl` です。**回答**、**Developer**、**トレース** をタブで切り替えます。さらに長いAnswerやTraceは、各TextBox内部のスクロールバーで内容だけを移動できます。

つまりスクロールは役割を分けています。

```text
右端のページスクロール
  → 画面全体を上下移動

回答 / トレース内部のスクロール
  → 長い本文だけを移動
```

通常起動時は最大化して表示します。タイトルバーの **最小化 / 元に戻す / 最大化** が使え、元に戻した後はウィンドウ端をドラッグして自由にサイズ変更できます。

`ResizeMode="CanResizeWithGrip"` を指定し、INPUTとRESULTの間には `GridSplitter` を配置しています。質問欄を広くしたい場合、または結果欄を広くしたい場合は境界を上下にドラッグできます。

送信時は「回答」タブを表示します。**エラー時も回答タブを維持**し、今回のQuestionとERROR内容を残します。Developer / トレースは利用者が必要に応じて開く診断領域とし、エラーだからといって自動で別タブへ移動しません。

## Question入力欄 / MVVM

Question欄はWPF標準の `TextBox` をそのまま編集用コントロールとして使います。

```text
ユーザーが QuestionBox へ入力
        ↓
QuestionBox.Text
        ↓
TextChanged で ViewModel.Question にも同期
        ↓
Send時は QuestionBox.Text をAPIへ送信
```

PowerShell + `XamlReader` では、主要な入力欄までData Bindingだけに依存すると環境差の切り分けが難しくなるため、**質問の実入力は見えているTextBoxそのものを正とする**実装にしています。ViewModelには状態把握用として同じ値を同期します。

`QuestionBox` は明示的に `IsReadOnly="False"`、`IsEnabled="True"`、`Focusable="True"`、`InputMethod.IsInputMethodEnabled="True"` としています。

また、API待機中もQuestion欄は無効化しません。応答を待ちながら次の質問を入力できます。

CIではWindows PowerShell 5.1上で、QuestionBoxの編集性、ResultTabsに「回答」「Developer」「トレース」の3タブがあること、Prompt example選択だけではQuestionを上書きしないこと、Mode変更でQuestionを上書きしないこと、エラー時に回答タブを維持してQuestion + ERRORを残すこと、Developerへ診断情報が入ることを確認します。さらに1200×900で各タブの表示領域、ページスクロール、`DataRowView` の `INotifyPropertyChanged`、Answer / Status のData Binding、Background Runspace自己テストも検証します。

参考: PowerShell / WPF / MVVMの考え方
- https://papanda925.com/?p=2187
- https://papanda925.com/?tag=xaml

## Modes

- `Chat` - 通常のChat Completions
- `Streaming` - `HttpClient` + `ResponseHeadersRead` でSSEを逐次処理
- `Tool Calling` - `calculate_sum` のTool Callをローカル実行して最終回答まで取得

Model欄は自由入力なので、無料ルーターと有料モデルで同じ実装を比較できます。

## 画面

1画面で次を確認できます。

- 質問
- 最大10往復の会話とOrcaRouterからの回答（「回答」タブ）
- HTTP Status / Elapsed / Token / Cost / Request / Response（「Developer」タブ）
- 処理ステップとHTTPトレース（「トレース」タブ）
- プロンプト例（任意）と明示的な「質問欄に挿入」

トレースには、APIキーを伏せたリクエスト情報、Request JSON、HTTPステータス、Response headers、Raw response、例外情報などを表示します。

## Quick start

Windows PowerShell 5.1 または Windows 上の PowerShell 7 で実行します。

このサンプルはWPF/XAMLを使うため、**`-STA` 付きで起動**してください。PowerShellの実行ポリシーで `.ps1` が止められる環境では、Windows PowerShell 5.1 の例のように、そのプロセスだけ `-ExecutionPolicy Bypass` を指定します。システム全体の実行ポリシーを恒久変更する必要はありません。

```powershell
cd PowerShell
powershell.exe -STA -ExecutionPolicy Bypass -File .\OrcaRouterChat.ps1
```

PowerShell 7 の場合:

```powershell
pwsh.exe -STA -File .\OrcaRouterChat.ps1
```

起動時のAPIキーはダミーです。

```text
xxx-your-orcarouter-api-key-xxx
```

画面上で実APIキーへ差し替えてください。既定モデルは `orcarouter/free` です。

APIキーがダミー値の間は、API Key欄の下に初回利用者向けリンクを表示します。

**[OrcaRouterを初めて利用する / APIキーを作成する](https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d)**

> 上記は本プロジェクトの紹介リンクです。紹介経由で登録されたWorkspaceの利用に応じて、開発者へ報酬が還元される場合があります。既にAPIキーを持っている場合は、そのキーをそのまま入力してください。

**Web版やVBA版へ入力したAPIキーはPowerShell版へ自動共有されません。** 起動したWPF画面のAPI Key欄へ個別に入力してください。

ローカル検証用に `OrcaRouterChat.ps1` へ一時的に埋め込む場所も明示しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

## 初回テストのおすすめ順

1. `Chat` で短い固定回答を試す
2. `Streaming` で回答が少しずつ表示されることを確認する
3. `Tool Calling` で `calculate_sum(123, 456)` を試す

Tool Callingはモデル対応状況やQuotaによって、ローカル関数を実行する前にAPI側から拒否される場合があります。Raw response / Trace の `HTTP Status` と `error.code` を確認してください。

モードの違いと各方式の比較は [はじめて使うときの手順](../docs/getting-started.md) を参照してください。

## Multi-turn Chat / New Chat / Prompt examples / Developer

WPF側で直近10往復の user / assistant 履歴を保持し、Background Runspaceへスナップショットを渡して次回の `messages` に再送します。**新しいチャット** はこのローカル履歴だけをクリアし、ModelやAPI Keyは残します。

「プロンプト例」で要約、初心者向け説明、コードレビュー、JSON、英訳を選び、**質問欄に挿入** を押すと雛形をQuestionへ入れられます。プルダウンを選択しただけでは質問欄を書き換えず、自動送信もしません。

送信直後は、未確定のQuestionを回答タブへ仮表示せず、**成功済み会話をそのまま維持してStatusだけを「送信中」にします**。Streamingでは実際にAssistantのdeltaを受信した時点から、今回Questionと途中回答を表示します。成功時に初めて今回Question + Assistantを履歴へ確定します。

APIエラー、timeout、入力検証エラーでは失敗turnを履歴へ追加しませんが、**今回Question + ERROR内容** を回答タブに一時表示して残します。回答タブを消したりトレースへ強制移動したりしません。Developerタブにも取得できたHTTP Status、Elapsed、Model、Request、Response/Error bodyを表示するため、原因を追いやすくしています。CostはOrcaRouterの `X-OrcaRouter-Include-Cost: true` を利用し、返らない場合は `(not returned)` と表示します。

Mode変更は通信方式の選択だけとし、入力中のQuestionを書き換えません。3実装共通の状態遷移ルールは [UI behavior contract](../docs/ui-behavior-contract.md) を参照してください。

## Common processing steps

Web版・VBA版と同じ6ステップにそろえています。

1. Validate inputs
2. Build request
3. Send HTTP POST
4. Receive response
5. Parse assistant message
6. Update UI and trace

コード内にも `STEP 1` ～ `STEP 6` の同じコメントを入れています。

詳細は [Common processing flow](../docs/processing-flow.md) を参照してください。

## PowerShell 5.1 / Runspace collection normalization

Background Runspaceへ渡した `History` は、空・単一要素・配列のどの形でも同じ意味になるよう、件数を `Get-HistoryTurnCount` で正規化して数えます。

Windows PowerShell 5.1 + `Set-StrictMode -Version Latest` では、Runspace境界を通った値に対してCollection型であることを前提に `.Count` を直接参照すると、値の形によって `PropertyNotFoundException` になる余地があります。そのため、`History` に対する直接の `.Count` は使用しません。

WorkerエラーにはDeveloper / Traceへ可能な範囲で次も残します。

- Exception Type
- Script line
- Position message
- Script stack trace

これにより、HTTP 200を受信した後のローカル処理エラーも、APIエラーと区別して追跡できます。

## Error handling / debug information

学習用として、通常のサンプルより診断情報を多めに残します。

- masked API key
- endpoint / model
- request body
- HTTP status
- response headers
- raw response body
- elapsed time
- exception type
- PowerShell script stack（ローカルのユーザー/リポジトリパスをプレースホルダー化）
- invocation position（ローカルのユーザー/リポジトリパスをプレースホルダー化）

実APIキーそのものはトレースに出力しません。エラーのStack/Positionにローカルパスが含まれる場合は、`<repository-root>` / `<USERPROFILE>` などへ置換してから画面へ表示します。

## API

```text
POST https://api.orcarouter.ai/v1/chat/completions
Authorization: Bearer <API_KEY>
Content-Type: application/json
```


## Advanced verification

Streamingでは `data:` SSE行を順に処理し、`choices[0].delta.content` を回答欄へ逐次反映します。Tool Callingでは `tools` / `tool_choice`、モデルの `tool_calls`、ローカル実行結果、2回目のAPI RequestをすべてTraceできます。

HTTPエラーは `error.code`、`error.type`、`Retry-After` 等を可能な範囲で診断します。

詳細: [Advanced API tests](../docs/advanced-features.md)

## Official OrcaRouter links

- Documentation: https://docs.orcarouter.ai/introduction
- Quickstart: https://docs.orcarouter.ai/getting-started/quickstart
- Start OrcaRouter (project referral link): https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d
- API key documentation: https://docs.orcarouter.ai/getting-started/get-api-key
- Models: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors

