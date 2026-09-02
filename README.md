# OrcaRouter Samples

**Learning-focused OrcaRouter API samples for Web, PowerShell/WPF, and Excel VBA — including Chat, Streaming, Tool Calling, request/response tracing, and error handling.**

[![Powered by OrcaRouter](https://img.shields.io/badge/Powered_by-OrcaRouter-2563eb)](https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d)
[![Validate samples](https://github.com/papanda925/OrcaRouter-Samples/actions/workflows/validate-samples.yml/badge.svg)](https://github.com/papanda925/OrcaRouter-Samples/actions/workflows/validate-samples.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

OrcaRouter API を **Web / PowerShell / Excel VBA** から呼び出す、学習者向けのサンプル集です。

3つの実装は、言語やUIが違っても処理を比較しやすいように、同じ6ステップ・同じ用語・同じTrace方針にそろえています。現在は **Chat / Streaming / Tool Calling** の3モードに加え、初回セットアップ、複数ターンChat、最大10往復の履歴、新しいチャット、プロンプト例、Developer Information を実装しています。

初めて動かす場合は、最初に **[はじめて使うときの手順](docs/getting-started.md)** を読んでください。VBAの `SetupOrcaRouterSample`、PowerShellの `-STA` / Execution Policy、Webローカルサーバー、APIキーの個別入力など、実際につまずきやすい点をまとめています。

## Project highlights

| 項目 | 内容 |
|---|---|
| Cross-environment | Web / PowerShell-WPF-XAML / Excel VBA の3方式を同じ考え方で比較 |
| OrcaRouter integration | `/v1/chat/completions` を直接利用し、既定モデルは `orcarouter/free` |
| API features | Chat / Streaming / Tool Calling |
| Observability | 成功/失敗の両方で HTTP Status / Elapsed / Token / Cost / Request JSON / Response JSON / HTTP Trace を確認可能 |
| PowerShell UI | Data Binding / `INotifyPropertyChanged` / Background Runspace |
| Safety | 実APIキーをソースへ保存しない設計、CIで秘密情報・個人パスを検査 |
| License | MIT License |

このリポジトリは、特定のフレームワークに隠さず、**OrcaRouter APIへ何を送信し、何が返り、各環境でどう処理するかを追えること**を重視しています。

## Samples

| Sample | UI | HTTP implementation | Status |
|---|---|---|---|
| [Web](Web/README.md) | HTML / CSS | Browser `fetch` / ReadableStream | Chat / Streaming / Tool Calling |
| [PowerShell](PowerShell/README.md) | WPF / XAML | Background Runspace + `System.Net.Http.HttpClient` | Chat / Streaming / Tool Calling |
| [VBA](VBA/README.md) | Excel worksheet / cells | `MSXML2.XMLHTTP.6.0` | Chat / Streaming / Tool Calling |

## First run checklist

| Sample | 最初に必要なこと | APIキー入力 |
|---|---|---|
| Web | `Web\start-server.ps1` でローカルサーバーを起動 | Web画面へ入力 |
| PowerShell | `-STA` で起動。必要に応じて `-ExecutionPolicy Bypass` | WPF画面へ入力 |
| VBA | 2つのBASをimport → compile → **`SetupOrcaRouterSample` を実行** | `OrcaRouter Chat` シート B3 |

**APIキーは3方式で自動共有されません。** Web / PowerShell / VBA のそれぞれへ個別に入力してください。

詳しい手順: [docs/getting-started.md](docs/getting-started.md)

### APIキーをまだ持っていない場合

初回セットアップでは、次の紹介リンクから OrcaRouter を開いてアカウント/APIキーを準備できます。

**[OrcaRouterを初めて利用する / APIキーを作成する](https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d)**

> このURLは本プロジェクトの紹介リンクです。紹介経由で登録されたWorkspaceの利用に応じて、プロジェクト開発者へ報酬が還元される場合があります。すでにOrcaRouterのAPIキーを持っている場合は、そのキーをそのまま各サンプルへ入力してください。

現在はこの紹介URLを設定しています。Built with OrcaRouter / Partner Dashboard 側のリポジトリ切替後にもURLを再確認し、必要があれば差し替えます。

## Priority 1 - beginner-friendly, developer-visible

通常画面はできるだけ簡単に保ち、APIの内部を見たい人だけ詳細へ進める構成です。

- **First run** - APIキーが未設定のときに、控えめな初回案内を表示
- **Multi-turn Chat** - アプリ側が会話履歴を保持し、次回の `messages` に再送
- **10-turn limit** - user + assistant を1往復として、直近10往復のみ保持
- **New chat** - APIへリセット電文は送らず、アプリ側の履歴だけをクリア
- **Prompt examples** - 選択だけではQuestionを書き換えず、明示的なInsert/Applyで要約、説明、コードレビュー、JSON、翻訳を挿入
- **Stable request state** - 送信中は回答待ち状態を明示し、成功時のみturnを内部履歴へcommit。回答欄には最新のAssistant回答だけを表示し、失敗時はERRORを表示
- **Developer Information** - 成功/失敗の両方でHTTP Status、Elapsed、Model、Prompt Tokens、Completion Tokens、Total Tokens、Cost、Request JSON、Response/Error body

Cost取得では OrcaRouter のChat Completions仕様にある `X-OrcaRouter-Include-Cost: true` を使用します。APIから `usage.cost_usd` が返らない場合は、推測せず `(not returned)` と表示します。

Web / PowerShell / VBA で共通に守る送信中・成功・失敗・New Chat・Prompt/Mode変更のルールを [UI behavior contract](docs/ui-behavior-contract.md) に明文化しています。CIもこの状態遷移契約を検査します。

## Chat / Streaming / Tool Calling

| Mode | 内容 | API呼び出しの特徴 |
|---|---|---|
| Chat | 完成した回答を1つのJSONで受信 | 基本のChat Completions |
| Streaming | SSEで回答を少しずつ受信 | `stream: true`、`delta.content` を順次処理 |
| Tool Calling | モデルがローカルToolを要求し、その結果を使って最終回答 | 通常2回のAPI呼び出し |

Tool CallingのサンプルToolは `calculate_sum(a, b)` です。モデルや無料ルーターの状況によっては、Tool実行前にQuotaやモデル対応の理由でAPI側から拒否されることがあります。

## OrcaRouter integration

このリポジトリでは、Web / PowerShell / VBA の各サンプルから OrcaRouter のChat Completions APIを直接呼び出します。

```text
POST https://api.orcarouter.ai/v1/chat/completions
Authorization: Bearer <API_KEY>
```

既定では `orcarouter/free` を使い、Model欄を変更して利用可能な別モデルでも検証できます。無料モデルだけでなく、同じOpenAI互換形式でStreamingやTool Calling、エラー処理を比較できることを目的としています。

このプロジェクトは **papanda925 が開発・保守する独立したオープンソースの学習用サンプル**です。

## Official OrcaRouter links

- Official site: https://www.orcarouter.ai/
- Japanese site: https://www.orcarouter.ai/ja
- Documentation: https://docs.orcarouter.ai/introduction
- Quickstart: https://docs.orcarouter.ai/getting-started/quickstart
- Start OrcaRouter (project referral link): https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d
- API key documentation: https://docs.orcarouter.ai/getting-started/get-api-key
- Models: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors

## What you can see

各サンプルは、同じ画面または同じ操作単位で次の内容を確認しやすいようにしています。

1. **質問入力と最大10往復の内部会話履歴（次回Request用）**
2. **OrcaRouterからの最新回答**
3. **Developer Information（HTTP Status / Elapsed / Token / Cost）**
4. **Request JSON / Response JSON**
5. **Raw response / Raw JSON**
6. **処理ステップ / HTTP電文のTrace**

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
│  ├─ api-key-setup.md
│  └─ ui-behavior-contract.md
├─ Web/
│  ├─ index.html
│  ├─ style.css
│  ├─ app.js
│  ├─ ui-contract.test.js
│  └─ README.md
├─ PowerShell/
│  ├─ MainWindow.xaml
│  ├─ OrcaRouterChat.ps1
│  ├─ OrcaRouterWorker.ps1
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
- WPF Data Binding / INotifyPropertyChanged
- PowerShell Runspaceによる非同期処理
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
- persistent / long-term conversation storage
- unit / integration tests
- organization-specific security controls
