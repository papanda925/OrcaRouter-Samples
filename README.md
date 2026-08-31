# OrcaRouter Samples

OrcaRouter API を **Web / PowerShell / Excel VBA** から呼び出す、学習者向けのサンプル集です。

3つの実装は、言語やUIが違っても処理を比較しやすいように、同じ6ステップ・同じ用語・同じTrace方針にそろえています。現在は **Chat / Streaming / Tool Calling** の3モードを実装しています。

## Samples

| Sample | UI | HTTP implementation | Status |
|---|---|---|---|
| [Web](Web/README.md) | HTML / CSS | Browser `fetch` / ReadableStream | Chat / Streaming / Tool Calling |
| [PowerShell](PowerShell/README.md) | WPF / XAML | `System.Net.Http.HttpClient` | Chat / Streaming / Tool Calling |
| [VBA](VBA/README.md) | Excel worksheet / cells | WinHTTP + curl(SSE) | Chat / Streaming / Tool Calling |

## What you can see

各サンプルは、同じ画面または同じ操作単位で次の3点を確認できるようにしています。

1. **質問**
2. **OrcaRouterからの回答**
3. **処理ステップ / HTTP電文のTrace**

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

このリポジトリに実APIキーをコミットしないでください。

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
