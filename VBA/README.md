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
| Raw JSON title/status | J1:P2 |
| Raw JSON body | J3:P15 |
| Trace | A18:E... |
| Send | 図形ボタン |
| Clear trace | 図形ボタン |

`OrcaRouterSample.bas` を標準モジュールとしてインポートし、`SetupOrcaRouterSample` を1回実行すると画面を自動生成します。

## Files

- `OrcaRouterSample.bas` - シートUI作成、通常Chat、共通helper
- `OrcaRouterAdvanced.bas` - Streaming、Tool Calling、高度なエラー診断

VBEの文字コード差異で壊れないよう、2つの `.bas` は **ASCIIのみ** で記述しています。コード内のコメント・UI文言・エラーメッセージは英語です。READMEは日本語のままです。

外部JSONライブラリを必須にしないため、この学習版では Chat Completions の `content` 文字列を取り出す軽量な処理を含めています。本番コードではフル機能のJSONパーサー利用を推奨します。

## Quick start

VBEの `.bas` インポートはWindowsのシステムコードページに依存し、日本語を含むソースは環境によって文字化けします。
そのため、このリポジトリのVBAソースは **ASCII-only** に変更しました。Shift-JIS変換は不要です。

1. PowerShellで最新の `main` を取得
2. 新しいExcelブックを `.xlsm` 形式で保存
3. `Alt + F11` でVBEを開く
4. 既に文字化けした `OrcaRouterSample` / `OrcaRouterAdvanced` がある場合は削除
5. `ファイル > ファイルのインポート` から `VBA\OrcaRouterSample.bas` と `VBA\OrcaRouterAdvanced.bas` を直接読み込む
6. `デバッグ > VBAProjectのコンパイル` を実行
7. `SetupOrcaRouterSample` を実行
8. 作成された `OrcaRouter Chat` シートの B3 にAPIキーを入力
9. 「Send」ボタンを押す

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

通常ChatとTool Callingでは、参照設定を追加せずに使える late binding の `MSXML2.XMLHTTP.6.0` を利用します。

### XMLHTTPへ変更した理由

PowerShell版は .NET の `HttpClient` で比較的すぐ応答する一方、VBA版の `WinHttp.WinHttpRequest.5.1` では、HTTP Statusを受け取る前の待機が長くなる実測がありました。

`WinHttp.WinHttpRequest` はWindowsのWinHTTPスタックを使い、主にマシン単位のネットワーク／プロキシ設定を参照します。一方、`MSXML2.XMLHTTP.6.0` はデスクトップアプリケーション向けで、現在のユーザー環境のネットワーク／プロキシ設定に近い経路を使います。

そのため、このExcel VBAサンプルでは通常ChatとTool Callingを `XMLHTTP` に切り替えています。PowerShellの `HttpClient` と完全に同じ実装ではありませんが、少なくともWinHTTP固有の通信経路を外して比較できるようにしています。

### 非同期処理にしている理由

`.Open "POST", URL, True` のように **第3引数を `True` にして非同期送信**します。

同期処理（`False`）では、`.Send` の中でHTTP処理が完了するまでVBAへ制御が戻りません。その間は `DoEvents` やTrace更新が実行できないため、API応答が遅いとExcelが固まったように見えます。

非同期処理（`True`）では、`.Send` のあとVBAへ制御が戻ります。XMLHTTPにはWinHTTPの `WaitForResponse` はないため、`readyState` を確認します。

```text
.Send
  ↓
readyState を確認
  ↓
4 ではない
  ↓
Trace / StatusBar 更新
  ↓
50ms待機
  ↓
YieldToExcel → DoEvents
  ↓
もう一度 readyState を確認
```

`readyState = 4` になればレスポンス完了です。

`YieldToExcel` と `WaitForXmlHttpResponse` はこのサンプルで用意した補助処理です。`DoEvents` を毎ループ無制限に呼ばず、Excelの再描画やウィンドウ操作に必要な範囲で利用します。また、待機中のSend二重実行を防ぐフラグも持たせています。

通常Chatでタイムアウトした場合は、まず `TestOrcaRouterConnection` マクロを実行してください。このマクロも `MSXML2.XMLHTTP.6.0` で `GET /v1/models` を呼び出し、次を切り分けます。

- 2xx: ネットワーク到達性とAPIキー認証は動作。Chat側のモデル選択・ルーティング・応答待ちを確認
- 401 / 403: APIキー、権限、Quotaなどを確認
- HTTP応答前のタイムアウト: ユーザー側ネットワーク、プロキシ、セキュリティ製品などを確認

```text
POST https://api.orcarouter.ai/v1/chat/completions
Authorization: Bearer <API_KEY>
Content-Type: application/json
```

## Answer / Raw JSON

APIからHTTPレスポンスを受信すると、同じレスポンスを2つの見方で表示します。

- **Answer（B11:H15）**: `choices[0].message.content` からユーザー向け回答本文を取り出して表示
- **Raw JSON（J3:P15）**: APIから受信したレスポンス本文を、解析前の文字列のまま表示

通常Chatでは受信したJSON全体をRaw JSONへ表示してからAnswerを解析します。そのため、Answer抽出でエラーになった場合でもRaw JSONを見ればAPIが実際に何を返したか確認できます。

Answer抽出は、レスポンス全体から単純に最初の `content` を探すのではなく、`choices` → `message` の位置を確認してから `content` を取得します。また、`content` が文字列ではなくテキストパーツ配列の場合は、最初の `text` をフォールバックとして使用します。

Tool Callingでは1回目のTool CallレスポンスをRaw JSONへ表示し、2回目の最終レスポンス受信後にRaw JSON欄を最終レスポンスへ更新します。Answerには2回目の `message.content` を表示します。

Streamingでは1つのJSONレスポンスではなくSSEイベントが連続するため、Raw JSON欄には最後に受信したJSON形式のSSEイベントを表示します。Answer欄には各SSEイベントの `delta.content` を連結した最終テキストを表示します。

Raw JSONはExcelセルの上限と可読性を考慮し、非常に長い場合は約30,000文字で切り詰めます。Traceには従来どおりHTTP Status、Headers、Raw response等も記録します。

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
- 通常ChatとTool Callingは `MSXML2.XMLHTTP.6.0` の非同期送信を使い、`readyState` を短い間隔で確認します。
- 待機ループ、Streamingの逐次更新、Trace描画では `DoEvents` を直接ばらまかず、`YieldToExcel` を通して約0.05秒間隔に抑制しています。
- 待機中はExcelのステータスバーへ経過秒数を表示し、一定間隔でTraceにも `WAIT` 行を追加します。
- 通常Chatの総待機上限は120秒です。
- 本番用途ではAPIキーをワークシートへ平文保存しない設計に変更してください。


## Modes

B5セルのドロップダウンで選択します。

- `Chat` - `MSXML2.XMLHTTP.6.0` による非同期Chat
- `Streaming` - Windows標準 `curl.exe` をVBAから起動し、SSEを標準出力から逐次読取
- `Tool Calling` - `calculate_sum(a, b)` のTool Callをローカル実行し、Tool結果を含む2回目のAPI呼び出しまで実施

Streaming Modeでは、APIキーをcurlのコマンドライン引数へ直接書かず、stdinから渡すcurl configにAuthorization headerを設定します。

Model欄は自由入力なので、無料ルーターから利用可能な有料モデルへ変更してAPI互換性を比較できます。

詳細: [Advanced API tests](../docs/advanced-features.md)


## VBE import encoding

VBEはUnicode対応のソースインポートを安定して扱えず、Windowsのシステムコードページによって結果が変わります。
そのためVBAの2つの `.bas` はASCII-onlyをCIで強制しています。これにより、日本語Windows・英語Windows・UTF-8システムロケールなどの差に影響されずインポートできます。
