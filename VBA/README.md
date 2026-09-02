# OrcaRouter VBA Sample

Excel VBA から OrcaRouter の Chat Completions API を呼び出し、**Chat / Streaming / Tool Calling** をExcelシート上で比較できる学習者向けサンプルです。

## Architecture

このサンプルは **ExcelシートをそのままUIとして使います**。

外部UserFormは使わず、セルと図形ボタンだけで次を表現します。

| Area | Excel |
|---|---|
| First-run API-key link | B2:H2 |
| API key | B3 |
| Model | B4 |
| Mode | B5:D5 |
| Question | B6:H9 |
| Prompt template (optional) | B10:D10 |
| Answer | B11:H15 |
| Status | B16:H16 |
| Raw JSON title/status | J1:P2 |
| Response JSON | J3:P15 |
| Developer metrics | J17:P20 |
| Request JSON | J22:P30 |
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
7. **`SetupOrcaRouterSample` を実行**
8. 作成された `OrcaRouter Chat` シートの B3 にAPIキーを入力
9. 必要に応じて `RunOrcaRouterVbaSelfTests` / `TestOrcaRouterConnection` を実行
10. 「Send」ボタンを押す

### Setupは最初に必須

2つのBASファイルをimportしただけでは、入力欄・Answer・Raw JSON・Trace・Sendボタンを持つサンプルシートは完成しません。

最初に `SetupOrcaRouterSample` を実行してください。

また、Setupを再実行するとサンプルシートを作り直すため、B3のAPIキーや入力値がダミー値へ戻ることがあります。**Setup後はB3を必ず再確認**してください。

**Web版やPowerShell版へ入力したAPIキーはExcelへ自動共有されません。** VBA版ではB3へ個別に入力します。

初期APIキーはダミーです。

```text
xxx-your-orcarouter-api-key-xxx
```

既定モデルは `orcarouter/free` です。

B3セルへの入力だけでなく、ローカル検証用に `OrcaRouterSample.bas` の定数へ一時的に埋め込む方法も用意しています。詳しくは [APIキーの設定方法](../docs/api-key-setup.md) を参照してください。

## 初回テストのおすすめ順

1. `Chat` で短い固定回答を確認
2. `Streaming` でAnswerが少しずつ増えることを確認
3. `Tool Calling` で `123 と 456 を足してください。` を確認

内部ロジックだけを確認する場合:

```text
RunOrcaRouterVbaSelfTests
```

API到達性と認証を切り分ける場合:

```text
TestOrcaRouterConnection
```

Tool Callingが失敗した場合は、まずRaw JSONの `HTTP Status` / `error.code` を確認してください。Quotaやモデル対応の問題では、`calculate_sum` を実行する前にAPI側で止まることがあります。

3方式共通の手順とモード比較は [はじめて使うときの手順](../docs/getting-started.md) を参照してください。

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

Chat / Streaming / Tool Calling は、参照設定を追加せずに使える late binding の `MSXML2.XMLHTTP.6.0` を利用します。

### XMLHTTPへ変更した理由

PowerShell版は .NET の `HttpClient` で比較的すぐ応答する一方、VBA版の `WinHttp.WinHttpRequest.5.1` では、HTTP Statusを受け取る前の待機が長くなる実測がありました。

`WinHttp.WinHttpRequest` はWindowsのWinHTTPスタックを使い、主にマシン単位のネットワーク／プロキシ設定を参照します。一方、`MSXML2.XMLHTTP.6.0` はデスクトップアプリケーション向けで、現在のユーザー環境のネットワーク／プロキシ設定に近い経路を使います。

そのため、このExcel VBAサンプルでは Chat / Streaming / Tool Calling を `XMLHTTP` に統一しています。PowerShellの `HttpClient` と完全に同じ実装ではありませんが、少なくともWinHTTP固有の通信経路や外部 `curl.exe` プロセスを使わず、Excel/VBA内で比較できるようにしています。

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
- 401 / 402 / 403: APIキー、権限、Quota、無料枠・残高などを確認
- HTTP応答前のタイムアウト: ユーザー側ネットワーク、プロキシ、セキュリティ製品などを確認

```text
POST https://api.orcarouter.ai/v1/chat/completions
Authorization: Bearer <API_KEY>
Content-Type: application/json
```

## Conversation / Raw JSON

APIからHTTPレスポンスを受信すると、同じ結果を利用者向け表示と開発者向けRaw JSONの2つの見方で確認できます。

- **Answer（B11:H15）**: 最新のAssistant回答、Streaming途中回答、またはERRORだけを表示
- **Raw JSON（J3:P15）**: APIから受信したレスポンス本文または最新SSEイベントを、解析前の内容として表示

通常Chatでは受信したJSON全体をRaw JSONへ表示してからAssistant本文を解析します。そのため、本文抽出でエラーになった場合でもRaw JSONを見ればAPIが実際に何を返したか確認できます。

Answer抽出は、レスポンス全体から単純に最初の `content` を探すのではなく、`choices` → `message` の位置を確認してから `content` を取得します。また、`content` が文字列ではなくテキストパーツ配列の場合は、最初の `text` をフォールバックとして使用します。

Tool Callingでは1回目のTool CallレスポンスをRaw JSONへ表示し、2回目の最終レスポンス受信後にRaw JSON欄を最終レスポンスへ更新します。Conversationには成功時に2回目の `message.content` をAssistantとして確定します。

Streamingでは1つのJSONレスポンスではなくSSEイベントが連続します。`XMLHTTP.responseText` を `readyState = 3 (LOADING)` の間から差分読取し、`data: {...}` 行を順次解析します。Raw JSON欄には最後に受信したJSON形式のSSEイベントを表示し、Conversationには各SSEイベントの `delta.content` を連結して今回Questionとともに逐次表示します。MSXMLがHTTPレスポンスをUnicode文字列へ変換した後にセルへ書くため、コンソール標準出力経由の文字化けを避けられます。

Raw JSONはExcelセルの上限と可読性を考慮し、非常に長い場合は約30,000文字で切り詰めます。Traceには従来どおりHTTP Status、Headers、Raw response等も記録します。

## Multi-turn Chat / New chat / Prompt template / Developer Information

Chat / Streaming / Tool Calling は、VBAモジュール内の**同じ成功済み会話履歴**を使います。直近10往復の user / assistant を保持し、次回の `messages` JSONへ再び含めます。シート上の **New chat** ボタンは履歴だけをクリアし、API Key / Model / Modeは残します。

B10:D10 は **Prompt template (optional)** です。最初は `(select)` で、テンプレートを選んだだけではQuestionは変わりません。選択後に **Insert prompt** を押したときだけ、要約・初心者向け説明・コードレビュー・JSON・翻訳の雛形をQuestionへ入れます。

Send直後はAnswerをクリアし、B16:H16のStatusへ待機中であることを表示します。通常Chat / Tool Callingは完了時に最新Assistant回答だけをAnswerへ表示し、Streamingは実際にdeltaを受信した時点から途中回答だけを表示します。

成功時だけ今回Question + Assistantを内部履歴へ確定します。APIエラー、timeout、入力検証エラーは履歴へ追加せず、AnswerにはERROR内容だけを表示します。Streaming途中で失敗した場合は、受信済みpartial answerにERRORを追記します。Question自体はB6:H9に残るため、Answerへ重複表示しません。

Developer Informationでは HTTP Status、Elapsed、Model、Prompt / Completion / Total Tokens、Costを確認できます。Request JSONはJ22:P30、Response JSON / Error bodyはRaw JSON領域J3:P15で確認します。Chat / Streaming / Tool Callingの成功・失敗の両方で、取得できた診断情報を更新します。Tool Callingでは2回のAPIレスポンスにUsage/Costがある場合、それらを合算します。Cost取得には `X-OrcaRouter-Include-Cost: true` を使い、APIが値を返さない場合は推測しません。

3実装共通の状態遷移ルールは [UI behavior contract](../docs/ui-behavior-contract.md) を参照してください。

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
- Chat / Streaming / Tool Calling は `MSXML2.XMLHTTP.6.0` の非同期送信を使い、`readyState` を短い間隔で確認します。
- 待機ループ、Streamingの逐次更新、Trace描画では `DoEvents` を直接ばらまかず、`YieldToExcel` を通して約0.05秒間隔に抑制しています。
- 待機中はExcelのステータスバーへ経過秒数を表示し、一定間隔でTraceにも `WAIT` 行を追加します。
- 通常Chatの総待機上限は120秒です。
- 本番用途ではAPIキーをワークシートへ平文保存しない設計に変更してください。


## Modes

B5セルのドロップダウンで選択します。

- `Chat` - `MSXML2.XMLHTTP.6.0` による非同期Chat
- `Streaming` - `MSXML2.XMLHTTP.6.0` の `responseText` を差分読取し、SSE `data:` 行をVBA内で逐次解析
- `Tool Calling` - `calculate_sum(a, b)` のTool Callをローカル実行し、Tool結果を含む2回目のAPI呼び出しまで実施

Streaming Modeでも外部PowerShellや `WScript.Shell`、`curl.exe` は使用しません。

Model欄は自由入力です。`orcarouter/free` が `free_quota_exhausted` を返した場合、このサンプルは有料モデルへ自動切替しません。意図しない課金を避けるため、利用者が残高・モデル料金を確認したうえで、B4へ具体的な有料モデルを明示的に入力してください。

詳細: [Advanced API tests](../docs/advanced-features.md)


## VBE import encoding

VBEはUnicode対応のソースインポートを安定して扱えず、Windowsのシステムコードページによって結果が変わります。
そのためVBAの2つの `.bas` はASCII-onlyをCIで強制しています。これにより、日本語Windows・英語Windows・UTF-8システムロケールなどの差に影響されずインポートできます。

## Official OrcaRouter links

- Documentation: https://docs.orcarouter.ai/introduction
- Quickstart: https://docs.orcarouter.ai/getting-started/quickstart
- Start OrcaRouter (project referral link): https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d
- API key documentation: https://docs.orcarouter.ai/getting-started/get-api-key
- Models: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors
