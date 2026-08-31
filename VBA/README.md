# OrcaRouter VBA Sample

Excel VBA から OrcaRouter の Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** をExcelシート上で比較できる学習者向けサンプルです。

## Architecture

このサンプルは **ExcelシートをそのままUIとして使います**。

外部UserFormは使わず、セルと図形ボタンだけで次を表現します。

| Area | Excel |
|---|---|
| API key | B3 |
| Model | B4 |
| Mode | B5:D5 |
| Question | B6:H9 |
| Answer | B11:H15 |
| Trace | A18:E... |
| Send | 図形ボタン |
| Clear trace | 図形ボタン |

`OrcaRouterSample.bas` を標準モジュールとしてインポートし、`SetupOrcaRouterSample` を1回実行すると画面を自動生成します。

## Files

- `OrcaRouterSample.bas` - シートUI作成、通常Chat、共通helper
- `OrcaRouterAdvanced.bas` - Streaming、Tool Calling、高度なエラー診断

外部JSONライブラリを必須にしないため、この学習版では Chat Completions の `content` 文字列を取り出す軽量な処理を含めています。本番コードではフル機能のJSONパーサー利用を推奨します。

## Quick start

1. 新しいExcelブックを `.xlsm` 形式で保存
2. `Alt + F11` でVBEを開く
3. `ファイル > ファイルのインポート` から `OrcaRouterSample.bas` と `OrcaRouterAdvanced.bas` を読み込む
4. `SetupOrcaRouterSample` を実行
5. 作成された `OrcaRouter Chat` シートの B3 にAPIキーを入力
6. 「送信」ボタンを押す

初期APIキーはダミーです。

```text
xxx-your-orcarouter-api-key-xxx
```

既定モデルは `orcarouter/free` です。

B3セルへの入力だけでなく、ローカル検証用に `OrcaRouterSample.bas` の定数へ一時的に埋め込む方法も用意しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

## Common processing steps

Web版・PowerShell版と同じ6ステップにそろえています。

1. Validate inputs
2. Build request
3. Send HTTP POST
4. Receive response
5. Parse assistant message
6. Update UI and trace

VBAコード内にも `STEP 1` ～ `STEP 6` の同じコメントを配置しています。

詳細は [Common processing flow](../docs/processing-flow.md) を参照してください。

## HTTP implementation

HTTP通信には、参照設定を追加せずに使える late binding の `WinHttp.WinHttpRequest.5.1` を利用します。

```text
POST https://api.orcarouter.ai/v1/chat/completions
Authorization: Bearer <API_KEY>
Content-Type: application/json
```

## Trace / debug

シート下部に次を時系列で記録します。

- Time
- Step
- Direction
- Detail
- masked API key
- endpoint / model
- request JSON
- HTTP status
- elapsed time
- response headers
- raw response body
- VBA error number / source / description

実APIキーそのものはTraceへ出力しません。

Excelセルの最大文字数と可読性を考慮し、非常に長いTraceデータは一定長で切り詰めます。

## Notes

- Windows版Excelを想定しています。
- API呼び出しは同期処理です。学習しやすさを優先し、非同期化はしていません。
- 本番用途ではAPIキーをワークシートへ平文保存しない設計に変更してください。


## Modes

B5セルのドロップダウンで選択します。

- `Chat` - `WinHttp.WinHttpRequest.5.1` による通常Chat
- `Streaming` - Windows標準 `curl.exe` をVBAから起動し、SSEを標準出力から逐次読取
- `Tool Calling` - `calculate_sum(a, b)` のTool Callをローカル実行し、Tool結果を含む2回目のAPI呼び出しまで実施

Streaming Modeでは、APIキーをcurlのコマンドライン引数へ直接書かず、stdinから渡すcurl configにAuthorization headerを設定します。

Model欄は自由入力なので、無料ルーターから利用可能な有料モデルへ変更してAPI互換性を比較できます。

詳細: [Advanced API tests](../docs/advanced-features.md)
