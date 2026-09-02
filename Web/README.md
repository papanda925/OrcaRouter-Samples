# OrcaRouter Web Sample

ブラウザから OrcaRouter の OpenAI 互換 Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** を切り替えて比較できる学習者向けサンプルです。

## Modes

- `Chat` - 通常のChat Completions
- `Streaming (SSE)` - `stream: true` のSSEを逐次表示
- `Tool Calling (calculate_sum)` - OpenAI-style `tools` / `tool_choice` の2段階呼び出し

Model欄は自由入力なので、`orcarouter/free` から利用可能な有料モデルIDへ変更して同じコードで互換性を比較できます。

## 画面

1画面で次を確認できます。

- 最大10往復の会話とOrcaRouterからの回答
- プロンプト例（任意）+ 明示的な「質問欄に挿入」
- Developer Information（HTTP Status / Elapsed / Token / Cost）
- Request JSON / Response JSON
- APIから返ったRaw JSON
- 処理ステップとHTTPトレース

Raw JSON欄はVBA版と同じ考え方で、Chatではレスポンス本文全体、Tool Callingでは1回目→最終2回目のレスポンス、Streamingでは最新のSSE JSONイベントを表示します。Answer解析に失敗した場合でもAPIが実際に返した内容を確認できます。

トレースには、APIキーを伏せたリクエスト情報、リクエストJSON、HTTPステータス、応答ヘッダー、Raw response、エラー情報などを表示します。

## Files

- `index.html` - 画面
- `style.css` - モダンでシンプルなUI
- `app.js` - API呼び出し、6ステップの処理、会話履歴、トレース、エラー処理
- `ui-contract.test.js` - Prompt / Mode / Send / Error表示の回帰テスト
- `start-server.ps1` - Python不要のWindows用ローカルHTTPサーバー

## Quick start

まずローカルHTTPサーバーを起動してからブラウザで開きます。HTMLファイルを直接ダブルクリックするより、`start-server.ps1` または `python -m http.server` を使う方法を推奨します。

APIキーの初期値はダミーです。APIキーが未設定の間だけ、画面上部に初回利用者向けの控えめな案内を表示します。

**[OrcaRouterを初めて利用する / APIキーを作成する](https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d)**

> 上記は本プロジェクトの紹介リンクです。紹介経由で登録されたWorkspaceの利用に応じて、開発者へ報酬が還元される場合があります。既にAPIキーを持っている場合は、そのキーをそのまま入力してください。



```text
xxx-your-orcarouter-api-key-xxx
```

OrcaRouterで発行したAPIキーへ画面上で差し替えてください。

**PowerShell版やVBA版へ入力したAPIキーはWeb版へ自動共有されません。** Web画面へ個別に入力してください。

OrcaRouter Console からAPIキーのファイルをダウンロードした場合は、Web画面の **「キーのファイルを読込」** を押して、そのファイルを選択できます。ブラウザがファイルをローカルで読み取り、`sk-orca-...` 形式の完全なキーをAPI Key欄へ設定します。**ファイル名、ファイル本文、完全なキーはTraceへ出力しません。** Traceにはローカルファイルが選択されたこと、サイズ、マスク済みキーだけを表示します。

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

## 初回テストのおすすめ順

1. `Chat` で短い固定回答を確認
2. `Streaming` で回答が逐次表示されることを確認
3. `Tool Calling` で `calculate_sum(123, 456)` を確認

Chatの初期質問は、モデル知識に依存しない短い動作確認用です。

```text
日本語で「こんにちは。Web版Chatのテストです。」とだけ答えてください。
```

StreamingではAnswerが少しずつ増え、Raw JSONには最新のSSE JSONイベントが表示されます。

Tool Callingでは、モデルや無料ルーターのQuota状況によってローカル関数実行前にAPI側から拒否されることがあります。その場合はRaw JSON / Trace の `HTTP Status` と `error.code` を確認してください。

3方式共通の手順とモード比較は [はじめて使うときの手順](../docs/getting-started.md) を参照してください。

## Multi-turn Chat / New Chat / Prompt examples / Developer Information

Chat / Streaming / Tool Calling は、API側へ会話を保存させるのではなく、ブラウザ内の配列に直近10往復の成功済み user / assistant を保持し、次回の `messages` へ再送します。**New chat** はこのローカル履歴だけを消し、API Key / Model / Modeは変更しません。

**プロンプト例（任意）** は、プルダウンを選択しただけではQuestionを書き換えません。例を選び **「質問欄に挿入」** を押したときだけ定型文をQuestionへ入れます。挿入後に自由に編集してから送信します。Modeを変更しても入力中のQuestionは上書きしません。

送信直後は、未確定のQuestionを会話欄へ一瞬だけ表示することはせず、成功済みの会話をそのまま残してStatusを「送信中」にします。Streamingだけは実際にAssistantのdeltaを受信し始めた時点から、今回Questionと途中回答を表示します。

成功時は今回のQuestion + Assistantを履歴へ確定します。APIエラー、timeout、JSON解析エラー、入力検証エラーは履歴へ確定しませんが、**今回Question + ERROR内容を会話欄へ残す**ため、結果が消えたように見えません。

Developer Information は折りたたみ式です。正常時だけでなくエラー時も、取得できた HTTP Status、Elapsed、Model、Request JSON、Response/Error body を残します。Token / Cost が返った場合はそれも表示します。Costは `X-OrcaRouter-Include-Cost: true` で取得を依頼し、APIが `usage.cost_usd` を返さない場合は推測せず `(not returned)` と表示します。

3実装共通の状態遷移ルールは [UI behavior contract](../docs/ui-behavior-contract.md) を参照してください。

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

## Official OrcaRouter links

- Documentation: https://docs.orcarouter.ai/introduction
- Quickstart: https://docs.orcarouter.ai/getting-started/quickstart
- Start OrcaRouter (project referral link): https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d
- API key documentation: https://docs.orcarouter.ai/getting-started/get-api-key
- Models: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors
