# OrcaRouter Web Sample

ブラウザから OrcaRouter の OpenAI 互換 Chat Completions API を呼び出す、学習者向けの最小サンプルです。

## 画面

1画面で次の3点を確認できます。

- 質問
- OrcaRouterからの回答
- 処理ステップとHTTPトレース

トレースには、APIキーを伏せたリクエスト情報、リクエストJSON、HTTPステータス、応答ヘッダー、Raw response、エラー情報などを表示します。

## Files

- `index.html` - 画面
- `style.css` - モダンでシンプルなUI
- `app.js` - API呼び出し、6ステップの処理、トレース、エラー処理

## Quick start

APIキーの初期値はダミーです。

```text
xxx-your-orcarouter-api-key-xxx
```

OrcaRouterで発行したAPIキーへ画面上で差し替えてください。

ローカル検証用に `Web/app.js` へ一時的に埋め込む方法も用意しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

既定モデルは、学習用途で試しやすい `orcarouter/free` です。

ローカルHTTPサーバーで開く例:

```bash
cd Web
python -m http.server 8000
```

ブラウザで `http://localhost:8000/` を開きます。

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
