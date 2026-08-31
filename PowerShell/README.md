# OrcaRouter PowerShell Sample

PowerShell + WPF/XAML で OrcaRouter の Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** を切り替えて比較できる学習者向けサンプルです。

## Architecture

- `MainWindow.xaml` - 画面定義
- `OrcaRouterChat.ps1` - API呼び出し、トレース、エラー処理

画面と処理を分離して、XAMLを読める人・PowerShellを読める人のどちらにも追いやすい構成にしています。

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

ローカル検証用に `OrcaRouterChat.ps1` へ一時的に埋め込む場所も明示しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

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
- PowerShell script stack
- invocation position

実APIキーそのものはトレースに出力しません。

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
