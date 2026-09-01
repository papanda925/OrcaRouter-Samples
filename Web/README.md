# OrcaRouter Web Sample

ブラウザから OrcaRouter の OpenAI 互換 Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** を切り替えて比較できる学習者向けサンプルです。

## Modes

- `Chat` - 通常のChat Completions
- `Streaming (SSE)` - `stream: true` のSSEを逐次表示
- `Tool Calling (calculate_sum)` - OpenAI-style `tools` / `tool_choice` の2段階呼び出し

Model欄は自由入力なので、`orcarouter/free` から利用可能な有料モデルIDへ変更して同じコードで互換性を比較できます。

## 画面

1画面で次の4点を確認できます。

- 質問
- OrcaRouterからの回答
- APIから返ったRaw JSON
- 処理ステップとHTTPトレース

Raw JSON欄はVBA版と同じ考え方で、Chatではレスポンス本文全体、Tool Callingでは1回目→最終2回目のレスポンス、Streamingでは最新のSSE JSONイベントを表示します。Answer解析に失敗した場合でもAPIが実際に返した内容を確認できます。

トレースには、APIキーを伏せたリクエスト情報、リクエストJSON、HTTPステータス、応答ヘッダー、Raw response、エラー情報などを表示します。

## Files

- `index.html` - 画面
- `style.css` - モダンでシンプルなUI
- `app.js` - API呼び出し、6ステップの処理、トレース、エラー処理
- `start-server.ps1` - Python不要のWindows用ローカルHTTPサーバー

## Quick start

APIキーの初期値はダミーです。

```text
xxx-your-orcarouter-api-key-xxx
```

OrcaRouterで発行したAPIキーへ画面上で差し替えてください。

OrcaRouter Console からAPIキーのファイルをダウンロードした場合は、Web画面の **「キーのファイルを読込」** を押して、そのファイルを選択できます。ブラウザがファイルをローカルで読み取り、`sk-orca-...` 形式の完全なキーをAPI Key欄へ設定します。ファイル本文や完全なキーはTraceへ出力しません。

ローカル検証用に `Web/app.js` へ一時的に埋め込む方法も用意しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

既定モデルは、学習用途で試しやすい `orcarouter/free` です。

ローカルHTTPサーバーで開く例です。

### Windows / PowerShellだけで起動する場合（Python不要）

```powershell
cd Web
powershell.exe -ExecutionPolicy Bypass -File .\start-server.ps1
```

または、すでにPowerShellを開いている場合:

```powershell
.\start-server.ps1
```

起動後、ブラウザで `http://localhost:8000/` を開きます。

停止は `Ctrl + C` です。

### Pythonがある場合

```bash
cd Web
python -m http.server 8000
```

## Important security note

このサンプルは、学習のためにブラウザからAPIへ直接 `fetch` します。

**実APIキーをHTMLやJavaScriptへ埋め込んだままWeb公開しないでください。**

ブラウザに渡した秘密情報は利用者から参照できます。本番用途では、APIキーをサーバー側で保管し、ブラウザからは自分のバックエンドを呼び出す構成にしてください。

また、API側のCORS設定によってはブラウザからの直接呼び出しが拒否されることがあります。その場合も本番向けにはサーバー側プロキシ方式を利用してください。

## Common processing steps

PowerShell版・VBA版と同じ6ステップにそろえています。

1. Validate inputs
2. Build request
3. Send HTTP POST
4. Receive response
5. Parse assistant message
6. Update UI and trace

詳細は [Common processing flow](../docs/processing-flow.md) を参照してください。

## API

```text
POST https://api.orcarouter.ai/v1/chat/completions
Authorization: Bearer <API_KEY>
Content-Type: application/json
```

Request example:

```json
{
  "model": "orcarouter/free",
  "messages": [
    {
      "role": "user",
      "content": "こんにちは"
    }
  ]
}
```


## Advanced verification

StreamingではBrowserの `fetch()` + `ReadableStream` を使い、`TextDecoder("utf-8")` で受信バイト列をUnicode文字列へ変換しながら、`data: {...}` と終端の `data: [DONE]` を逐次処理します。SSE途中で `error` chunkが返された場合もエラーとして扱います。

Tool Callingではローカル関数 `calculate_sum(a, b)` を公開し、モデルからの `tool_calls` → ローカル実行 → `role: "tool"` を含む2回目のAPI呼び出しまでをTraceできます。

401 / 402 / 403 / 425 / 429 / 503などはHTTP Statusだけでなく、可能な範囲で `error.type`、`error.code`、`Retry-After` も表示します。

実APIで確認された `free_quota_exhausted` はHTTP Statusだけに依存せず `error.code` を優先して判定します。無料ルーターで拒否された場合に有料モデルへ自動切替はしません。

通常Chat / Tool Callingの待機上限は120秒、Streamingは90秒です。

詳細: [Advanced API tests](../docs/advanced-features.md)
