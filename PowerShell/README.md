# OrcaRouter PowerShell Sample

PowerShell + WPF/XAML で OrcaRouter の Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** を切り替えて比較できる学習者向けサンプルです。

## Architecture

- `MainWindow.xaml` - View（画面定義）
- `OrcaRouterChat.ps1` - ViewModel、API呼び出し、トレース、エラー処理

Model / Mode / Answer / Status は、`INotifyPropertyChanged` を実装した `OrcaRouterViewModel` とXAMLのData Bindingで連携します。Questionはユーザーが直接操作するWPF `TextBox` を入力元とし、`TextChanged` でViewModelへ同期します。

```text
View (XAML)
   ⇅ TwoWay / OneWay Binding
ViewModel (INotifyPropertyChanged)
   ↓
OrcaRouter API処理
```

これは **MVVMの考え方を取り入れた軽量構成**です。Send / ClearなどのイベントとQuestionの実入力はPowerShell/WPFコントロール側で扱うため、`ICommand` まで含めた厳密なフルMVVMではありません。

API KeyはWPFの `PasswordBox.Password` が標準では通常のData Binding対象にならないため、従来どおり直接取得します。Traceも大量追記を分かりやすくするため `AppendText()` を使用します。

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

CIではWindows PowerShell 5.1上で、QuestionBoxが **編集可能・有効・フォーカス可能・複数行向けの高さを持つこと**、Text変更がViewModelへ同期されることを確認します。

参考: PowerShell / WPF / MVVMの考え方
- https://papanda925.com/?p=2187
- https://papanda925.com/?tag=xaml

## Modes

- `Chat` - 通常のChat Completions
- `Streaming` - `HttpClient` + `ResponseHeadersRead` でSSEを逐次処理
- `Tool Calling` - `calculate_sum` のTool Callをローカル実行して最終回答まで取得

Model欄は自由入力なので、無料ルーターと有料モデルで同じ実装を比較できます。

## 画面

1画面で次の3点を確認できます。

- 質問
- OrcaRouterからの回答
- 処理ステップとHTTPトレース

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

**Web版やVBA版へ入力したAPIキーはPowerShell版へ自動共有されません。** 起動したWPF画面のAPI Key欄へ個別に入力してください。

ローカル検証用に `OrcaRouterChat.ps1` へ一時的に埋め込む場所も明示しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

## 初回テストのおすすめ順

1. `Chat` で短い固定回答を試す
2. `Streaming` で回答が少しずつ表示されることを確認する
3. `Tool Calling` で `calculate_sum(123, 456)` を試す

Tool Callingはモデル対応状況やQuotaによって、ローカル関数を実行する前にAPI側から拒否される場合があります。Raw response / Trace の `HTTP Status` と `error.code` を確認してください。

モードの違いと各方式の比較は [はじめて使うときの手順](../docs/getting-started.md) を参照してください。

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
- Get an API key: https://docs.orcarouter.ai/getting-started/get-api-key
- Models: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors

