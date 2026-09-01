# OrcaRouter Samples

OrcaRouter API を **Web / PowerShell / Excel VBA** から呼び出す、学習者向けのサンプル集です。

3つの実装は、言語やUIが違っても処理を比較しやすいように、同じ6ステップ・同じ用語・同じTrace方針にそろえています。現在は **Chat / Streaming / Tool Calling** の3モードを実装しています。

初めて動かす場合は、最初に **[はじめて使うときの手順](docs/getting-started.md)** を読んでください。VBAの `SetupOrcaRouterSample`、PowerShellの `-STA` / Execution Policy、Webローカルサーバー、APIキーの個別入力など、実際につまずきやすい点をまとめています。

## Samples

| Sample | UI | HTTP implementation | Status |
|---|---|---|---|
| [Web](Web/README.md) | HTML / CSS | Browser `fetch` / ReadableStream | Chat / Streaming / Tool Calling |
| [PowerShell](PowerShell/README.md) | WPF / XAML | `System.Net.Http.HttpClient` | Chat / Streaming / Tool Calling |
| [VBA](VBA/README.md) | Excel worksheet / cells | `MSXML2.XMLHTTP.6.0` | Chat / Streaming / Tool Calling |

## First run checklist

| Sample | 最初に必要なこと | APIキー入力 |
|---|---|---|
| Web | `Web\start-server.ps1` でローカルサーバーを起動 | Web画面へ入力 |
| PowerShell | `-STA` で起動。必要に応じて `-ExecutionPolicy Bypass` | WPF画面へ入力 |
| VBA | 2つのBASをimport → compile → **`SetupOrcaRouterSample` を実行** | `OrcaRouter Chat` シート B3 |

**APIキーは3方式で自動共有されません。** Web / PowerShell / VBA のそれぞれへ個別に入力してください。

詳しい手順: [docs/getting-started.md](docs/getting-started.md)

## Chat / Streaming / Tool Calling

| Mode | 内容 | API呼び出しの特徴 |
|---|---|---|
| Chat | 完成した回答を1つのJSONで受信 | 基本のChat Completions |
| Streaming | SSEで回答を少しずつ受信 | `stream: true`、`delta.content` を順次処理 |
| Tool Calling | モデルがローカルToolを要求し、その結果を使って最終回答 | 通常2回のAPI呼び出し |

Tool CallingのサンプルToolは `calculate_sum(a, b)` です。モデルや無料ルーターの状況によっては、Tool実行前にQuotaやモデル対応の理由でAPI側から拒否されることがあります。

## Official OrcaRouter links

- Official site: https://www.orcarouter.ai/
- Japanese site: https://www.orcarouter.ai/ja
- Documentation: https://docs.orcarouter.ai/introduction
- Quickstart: https://docs.orcarouter.ai/getting-started/quickstart
- Get an API key: https://docs.orcarouter.ai/getting-started/get-api-key
- Models: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors

## What you can see

各サンプルは、同じ画面または同じ操作単位で次の内容を確認しやすいようにしています。

1. **質問**
2. **OrcaRouterからの回答**
3. **Raw response / Raw JSON**
4. **処理ステップ / HTTP電文のTrace**

Traceには学習用として、通常より多めのデバッグ情報を残します。

- masked API key
- endpoint / model
- request headers
- request JSON
- HTTP status
- response headers where available
- raw response body
- elapsed time
- parse result
- runtime-specific error information

**実APIキーそのものはTraceへ出力しません。**

## Common six-step flow

3つのサンプルはコードコメントも同じステップ名に統一しています。

1. **STEP 1 - Validate inputs**
2. **STEP 2 - Build request**
3. **STEP 3 - Send HTTP POST**
4. **STEP 4 - Receive response**
5. **STEP 5 - Parse assistant message**
6. **STEP 6 - Update UI and trace**

共通シーケンス図とAPI電文は [docs/processing-flow.md](docs/processing-flow.md) を参照してください。Streaming / Tool Calling / 有料モデル互換性 / エラー処理の詳細は [docs/advanced-features.md](docs/advanced-features.md) にまとめています。

## API configuration

Endpoint:

```text
POST https://api.orcarouter.ai/v1/chat/completions
```

Authentication:

```text
Authorization: Bearer <API_KEY>
```

各サンプルのAPIキー初期値はダミーです。

```text
xxx-your-orcarouter-api-key-xxx
```

実行するときだけ、自分のOrcaRouter APIキーへ差し替えてください。

**どのファイルのどこを書き換えるか、ローカル限定でソースへ埋め込む方法、push前の確認方法まで含めた手順は [APIキーの設定方法](docs/api-key-setup.md) を参照してください。**

既定モデルは次の値にしています。

```text
orcarouter/free
```

モデル欄は変更可能なので、利用可能な別モデルへ差し替えることもできます。

## Minimal request

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

## Repository structure

```text
OrcaRouter-Samples/
├─ docs/
│  ├─ getting-started.md
│  ├─ processing-flow.md
│  ├─ advanced-features.md
│  └─ api-key-setup.md
├─ Web/
│  ├─ index.html
│  ├─ style.css
│  ├─ app.js
│  └─ README.md
├─ PowerShell/
│  ├─ MainWindow.xaml
│  ├─ OrcaRouterChat.ps1
│  └─ README.md
├─ VBA/
│  ├─ OrcaRouterSample.bas
│  ├─ OrcaRouterAdvanced.bas
│  └─ README.md
├─ LICENSE
└─ README.md
```

## Learning goals

このリポジトリでは、高度なフレームワークや抽象化よりも「HTTP APIを呼び出すとき実際に何が起きているか」を追いやすくすることを優先しています。

特に次を比較して学べます。

- JSON request の作り方
- Bearer認証
- HTTP POST
- OpenAI互換レスポンス
- UIとAPI処理の分離
- request / response trace
- エラー処理
- 実行環境ごとの違い

## Security notes

このリポジトリに実APIキーをコミットしないでください。APIキー、個人名、メールアドレス、PC固有の絶対パスなどを README / Issue / Trace の例へ貼る場合も、`<API_KEY>`、`<USER>`、`<repository-root>` などへダミー化してください。

特にWeb版は、学習のためブラウザからAPIを直接呼び出す構成です。ブラウザへ渡した秘密情報は利用者から参照できるため、実APIキーを埋め込んだ状態で公開してはいけません。本番WebアプリではAPIキーをサーバー側へ置いてください。

VBA版も学習しやすさを優先してシートからAPIキーを読み取ります。本番用途ではWindows Credential Manager、環境変数、社内のSecret管理基盤など、用途に合った方式へ置き換えてください。

## Scope

これらは学習・検証用サンプルです。

本番システムへ展開するときは、必要に応じて次を追加してください。

- secret management
- retry / exponential backoff
- rate-limit handling
- structured logging
- proxy / backend architecture
- conversation history
- unit / integration tests
- organization-specific security controls
