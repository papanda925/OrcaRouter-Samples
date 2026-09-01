# はじめて使うときの手順

このページは、OrcaRouter-Samples を **Web / PowerShell / Excel VBA の3方式で実際に動かすところまで**をまとめた手順です。

最初につまずきやすい点を先にまとめると、次の4点です。

1. **APIキーは完全な文字列が必要**です。画面上で一部が伏せられたキーは認証には使えません。
2. **APIキーは Web / PowerShell / VBA で共有されません。** それぞれの画面へ個別に入力します。
3. **VBA版は `SetupOrcaRouterSample` の実行が最初に必要**です。
4. **PowerShell版やWeb版のローカルサーバーは、環境によってPowerShellの実行ポリシーに止められることがあります。** その場合は、このページに記載した `-ExecutionPolicy Bypass` の起動方法を使います。

> このリポジトリは学習・検証用です。実APIキー、個人名、メールアドレス、PC固有の絶対パスをGitHubへ保存しないでください。

---

## 1. APIキーをまだ持っていない場合

初めてOrcaRouterを利用する場合は、次のリンクからアカウント/APIキーを準備できます。

**[OrcaRouterを初めて利用する / APIキーを作成する](https://www.orcarouter.ai/ref/ref_5074f764e512c8dd3d9d)**

> 上記は本プロジェクトの紹介リンクです。紹介経由で登録されたWorkspaceの利用に応じて、プロジェクト開発者へ報酬が還元される場合があります。すでにOrcaRouterのAPIキーを持っている場合は、新しく登録する必要はありません。既存キーをそのまま利用してください。

現在はこのURLを仮設定しています。Built with OrcaRouter / Partner Dashboard のリポジトリ切替完了後にも紹介URLを再確認し、必要があれば更新します。

## 2. OrcaRouter公式リンク

- 公式サイト: https://www.orcarouter.ai/
- 日本語サイト: https://www.orcarouter.ai/ja
- 公式ドキュメント: https://docs.orcarouter.ai/introduction
- Quickstart: https://docs.orcarouter.ai/getting-started/quickstart
- APIキー作成: https://docs.orcarouter.ai/getting-started/get-api-key
- モデル一覧: https://www.orcarouter.ai/models
- Streaming: https://docs.orcarouter.ai/advanced/streaming
- Tool Calling: https://docs.orcarouter.ai/advanced/tool-calling
- Errors: https://docs.orcarouter.ai/operations/errors

OrcaRouterのOpenAI互換APIのベースURLは次です。

```text
https://api.orcarouter.ai/v1
```

このサンプルでは Chat Completions を使います。

```text
POST https://api.orcarouter.ai/v1/chat/completions
```

---

## 3. APIキーを用意する

OrcaRouterでAPIキーを作成します。

APIキーは `sk-orca-` で始まる完全な文字列です。

このリポジトリでは、実キーの代わりに次のダミー値を使っています。

```text
xxx-your-orcarouter-api-key-xxx
```

### 重要

OrcaRouterの画面で既存キーが一部伏せ字になっている場合、その表示だけではAPI認証に使えません。

完全なキーを保存していない場合は、新しいキーを作成してください。

**完全なAPIキーを README、Issue、Chat、スクリーンショットへ貼る必要はありません。**

詳しい設定方法は [APIキーの設定方法](api-key-setup.md) を参照してください。

---

## 4. まずは最新ファイルを取得する

既にリポジトリをclone済みなら、リポジトリのルートで次を実行します。

```powershell
git pull origin main
```

以降の説明では、リポジトリを置いた場所を `<repository-root>` と表記します。

個人のユーザー名を含む絶対パスをREADMEやIssueへ記載する必要はありません。

---

# 5. Web版

## 起動

```powershell
cd <repository-root>\Web
.\start-server.ps1
```

PowerShellの実行ポリシーで止められた場合:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\start-server.ps1
```

起動後、ブラウザで次を開きます。

```text
http://localhost:8000/
```

停止は `Ctrl + C` です。

## APIキー

画面上部の **API key** に自分の完全なAPIキーを入力します。

または、OrcaRouter Consoleから保存したキーのテキストファイルを **「キーのファイルを読込」** から読み込めます。

APIキーはブラウザ内で使用されます。GitHubへ保存されません。

## 最初の確認

まず Mode を `Chat` にして送信します。

初期テストでは、モデル自身の知識に依存しない短い質問を推奨します。

```text
日本語で「こんにちは。Web版Chatのテストです。」とだけ答えてください。
```

正常時は次を確認できます。

- 会話欄にUser / Assistantが表示される
- 2回目以降のChatでは、直近の会話履歴がRequest JSONの `messages` に入る
- 履歴は最大10往復で、**新しいチャット** からクリアできる
- Developer Information に HTTP Status / Elapsed / Token / Cost が表示される
- Request JSON / Response JSON を確認できる
- Raw JSON にAPIレスポンスが表示される
- Trace に STEP 1 ～ STEP 6 が表示される
- HTTP Status が 2xx になる

### Web版の注意

Web版はブラウザからOrcaRouterへ直接 `fetch()` します。

公開Webサイトへ自分の実APIキーを埋め込んではいけません。公開用途では利用者自身にキーを入力してもらう方式か、サーバー側でAPIキーを管理するバックエンドプロキシ方式に変更してください。

---

# 6. PowerShell版

## 起動

Windows PowerShell 5.1 の例:

```powershell
cd <repository-root>\PowerShell
powershell.exe -STA -ExecutionPolicy Bypass -File .\OrcaRouterChat.ps1
```

PowerShell 7 の例:

```powershell
cd <repository-root>\PowerShell
pwsh.exe -STA -File .\OrcaRouterChat.ps1
```

### なぜ `-STA` が必要か

このサンプルは WPF / XAML の画面を使うため、STA（Single-Threaded Apartment）での起動を前提にしています。

### なぜ `-ExecutionPolicy Bypass` を付ける例があるか

Windowsの設定によっては、ローカルの `.ps1` 実行がExecution Policyで止められることがあります。

この例は、その1回のPowerShellプロセスでサンプルを起動するための指定です。システム全体の実行ポリシーを恒久的に変更する必要はありません。

### 日本語が文字化けしてParserErrorになる場合

Windows PowerShell 5.1は、BOMのないUTF-8スクリプトを誤った文字コードで読むことがあります。

```text
API繧ｭ繝ｼ...
UnexpectedToken
ParserError
```

のような表示になった場合は、まず `git pull origin main` で最新版へ更新してください。

このリポジトリの `PowerShell\OrcaRouterChat.ps1` は **UTF-8 BOM付き** にしてあり、Windows PowerShell 5.1の `powershell.exe -File` でも正しく日本語を読み込めるようにしています。

## APIキー

起動した画面の **API Key** に自分の完全なAPIキーを入力します。

Web版に入力したキーがPowerShell版へ自動的に引き継がれることはありません。

## 会話・Developer Information・トレースの見方

PowerShell版は **画面全体の縦スクロール + RESULTタブ内部のスクロール** の2段構成です。

```text
PageScrollViewer
  ├─ INPUT
  └─ RESULT
       ├─ 回答
       ├─ Developer
       └─ トレース
```

ウィンドウの高さが足りない場合は、右側にページ全体用の縦スクロールバーが自動表示され、下部まで移動できます。「回答」「Developer」「トレース」はタブで切り替え、長い回答やHTTPトレースは各タブ内部のスクロールバーで内容だけを移動します。

送信すると「回答」タブを表示し、エラーが発生した場合は「トレース」タブへ自動で切り替わります。INPUTとRESULTの境界はマウスで上下にドラッグして高さを変更できます。

## PowerShell版が画面を固めない仕組み

PowerShell版では、WPF画面とOrcaRouter API処理を分離しています。

```text
WPF UI
  ↓
Background PowerShell Runspace
  ↓
OrcaRouter API
  ↓
ConcurrentQueue
  ↓
ViewModel.Answer
  ↓
INotifyPropertyChanged
  ↓
Data Binding
  ↓
回答欄を更新
```

API待機はBackground Runspaceで行うため、通常ChatやTool Callingの応答待ち中もWPFのUIスレッドをHTTP待機で塞ぎません。Streamingも同じWorker側でSSEを読みます。

ViewModelにはC#クラスを埋め込まず、PowerShellで1行の `DataTable` を作り、その `System.Data.DataRowView` を使用します。`DataRowView` は `INotifyPropertyChanged` を実装しているため、AnswerやStatusの変更をData Bindingへ通知できます。

PowerShell版のファイルは次の3つです。

```text
PowerShell\MainWindow.xaml
PowerShell\OrcaRouterChat.ps1
PowerShell\OrcaRouterWorker.ps1
```


---

# 7. Excel VBA版

VBA版は、3方式の中で最初の準備が最も重要です。

## 1. Excelブックを準備

新しいExcelブックを作成し、マクロ有効ブックとして保存します。

```text
.xlsm
```

## 2. 2つのBASファイルをインポート

VBE（`Alt + F11`）を開き、標準モジュールとして次の2ファイルを読み込みます。

```text
VBA\OrcaRouterSample.bas
VBA\OrcaRouterAdvanced.bas
```

役割は次のとおりです。

| ファイル | 主な役割 |
|---|---|
| `OrcaRouterSample.bas` | UI作成、Chat、共通HTTP/JSON/Trace処理 |
| `OrcaRouterAdvanced.bas` | Streaming、Tool Calling、高度なエラー処理 |

## 3. コンパイル

VBEで次を実行します。

```text
デバッグ
→ VBAProjectのコンパイル
```

## 4. Setupを最初に実行

`Alt + F8` から次を実行します。

```text
SetupOrcaRouterSample
```

これで `OrcaRouter Chat` シートと、入力欄・Conversation・Prompt example・Developer Information・Request/Response JSON・Trace・Send/New chatボタンが作成されます。

**BASファイルをインポートしただけでは、操作用シートは完成しません。最初にSetupを実行してください。**

### Setupを再実行するときの注意

`SetupOrcaRouterSample` はサンプルシートを作り直します。

そのため、B3へ入力していたAPIキーやシート上の入力値はダミー値へ戻ることがあります。

Setup後は必ずAPIキーを再確認してください。

## 5. APIキーを入力

`OrcaRouter Chat` シートの次のセルです。

```text
B3 = API Key
B4 = Model
B5 = Mode
```

B3のダミー値を自分の完全なAPIキーへ変更します。

PowerShell版やWeb版へ入力したキーがExcelへ自動的に入ることはありません。

## 6. 自己テスト

API通信の前に、内部処理の確認として次を実行できます。

```text
RunOrcaRouterVbaSelfTests
```

接続だけを切り分けたい場合は次を実行します。

```text
TestOrcaRouterConnection
```

## 7. Send

Modeを選択してQuestionを入力し、**Send** を押します。

---

# Priority 1 の使い方

## 会話履歴

`/v1/chat/completions` へリセット電文を送るのではなく、アプリ側が会話を保持します。通常のChatでは次回Requestの `messages` へ、直近10往復の user / assistant を再送します。

```text
1回目: user
2回目: user + assistant + user
...
最大10往復
```

**New chat / 新しいチャット** はこのローカル履歴を空にします。API KeyとModelは消しません。

## Prompt example

要約、初心者向け説明、コードレビュー、JSON、翻訳の例を用意しています。例をそのまま送るのではなく、必要な文章やコードへ書き換えて使います。

## Developer Information

通常画面を複雑にしすぎないため、Webは折りたたみ、PowerShellはDeveloperタブ、VBAは右側のDeveloper領域として分離しています。

確認できる主な項目:

- HTTP Status
- Elapsed
- Model
- Prompt Tokens
- Completion Tokens
- Total Tokens
- Cost
- Request JSON
- Response JSON

Cost取得ではOrcaRouterの `X-OrcaRouter-Include-Cost: true` を利用します。APIが `usage.cost_usd` を返さない場合は金額を推測しません。

---

# 8. Chat / Streaming / Tool Calling の違い

3モードは同じ「質問を送る」機能に見えますが、APIの使い方が異なります。

| Mode | 何をするか | API呼び出し | 見どころ |
|---|---|---:|---|
| Chat | 回答が完成してから受け取る | 1回 | 最も基本的なChat Completions |
| Streaming | 回答をSSEで少しずつ受け取る | 1回 | `delta.content` が順次増える |
| Tool Calling | AIがローカル関数を選び、その結果を使って最終回答を作る | 通常2回 | AI → Tool → AI の往復 |

## Chat

基本形です。

```text
Question
  ↓
API request
  ↓
1つのJSON response
  ↓
Answer
```

最初の動作確認はChatから始めることを推奨します。

## Streaming

Requestに次を追加します。

```json
{
  "stream": true
}
```

APIからSSE形式で複数の `data: {...}` が返り、最後は `data: [DONE]` になります。

```text
Question
  ↓
API request (stream:true)
  ↓
data: {...}
data: {...}
data: {...}
  ↓
[DONE]
```

Web版は `fetch() + ReadableStream`、PowerShell版は `HttpClient + ResponseHeadersRead`、VBA版は `MSXML2.XMLHTTP.6.0` の `responseText` 差分読取で実装しています。

## Tool Calling

このサンプルでは `calculate_sum(a, b)` というローカルToolを用意しています。

テスト例:

```text
123 と 456 を足してください。
```

概念的な流れ:

```text
1回目のAPI
  ↓
modelが calculate_sum を要求
  ↓
ローカルで 123 + 456 = 579
  ↓
Tool結果を付けて2回目のAPI
  ↓
最終回答
```

Tool Calling対応可否はモデルに依存します。

また `orcarouter/free` が `free_quota_exhausted` などを返した場合は、Tool実行まで到達する前にAPI側で止まることがあります。

このサンプルは、意図しない課金を避けるため有料モデルへ自動切替しません。

---

# 9. 3方式のHTTP実装比較

| Sample | HTTP実装 | Streaming |
|---|---|---|
| Web | Browser `fetch()` | `ReadableStream` + `TextDecoder("utf-8")` |
| PowerShell | `.NET HttpClient` | `ResponseHeadersRead` + StreamReader |
| VBA | `MSXML2.XMLHTTP.6.0` | `readyState = 3` の `responseText` を差分読取 |

同じOrcaRouter APIを異なる実行環境から呼び出すことで、HTTP、JSON、非同期処理、Streaming、Tool Callingの違いを比較できます。

---

# 10. うまく動かないとき

## APIキーエラー

- ダミー値のままではないか
- 完全なAPIキーを入力しているか
- キーを作り直した場合、古いキーを使っていないか
- Web / PowerShell / VBA のそれぞれに正しいキーを入力したか

## `free_quota_exhausted`

無料ルーターの利用可能枠や無料モデルの状況によって発生します。

Tool Callingのコード自体が失敗したとは限りません。

Raw JSONと `error.code` を確認してください。

## VBAが動かない

- `.xlsm` で保存しているか
- 2つのBASをインポートしたか
- VBAProjectをコンパイルしたか
- `SetupOrcaRouterSample` を実行したか
- Setup後にB3のAPIキーを再入力したか

## PowerShell画面が起動しない

- `-STA` を付けているか
- Execution Policyで止められていないか
- Windows上で実行しているか

## Web版が開かない

- `start-server.ps1` が動いているか
- `http://localhost:8000/` を開いているか
- 8000番ポートを他のアプリが使用していないか

---

# 11. 公開前の安全確認

このリポジトリでは、サンプルのソースに実APIキーを保存しない方針です。

公開・Issue作成・スクリーンショット共有の前に、少なくとも次を確認してください。

- 実APIキーが写っていない
- APIキーを含むファイルを添付していない
- 個人名、メールアドレス、電話番号などが含まれていない
- `C:\Users\<USER>\...` のような個人PC固有のパスをそのまま掲載していない
- `/home/<USER>/...` のような個人環境のパスをそのまま掲載していない
- TraceやRaw JSONに秘密情報が含まれていない

READMEやIssueでは、ローカルパスは次のようなダミー表記を使います。

```text
<repository-root>
<USER>
<API_KEY>
xxx-your-orcarouter-api-key-xxx
```

一度外部へ公開したAPIキーは、文字列を後から削除するだけでは不十分です。無効化して新しいキーへ交換してください。
