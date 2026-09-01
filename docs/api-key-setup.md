# APIキーの設定方法

このリポジトリでは、実APIキーをGitHubへ保存しないよう、初期値を次のダミー値にしています。

```text
xxx-your-orcarouter-api-key-xxx
```

OrcaRouterで発行した実APIキーを使う方法は、次の2通りです。

**Web / PowerShell / VBA のAPIキー入力欄は独立しています。** 1つのサンプルへ入力したキーが、別のサンプルへ自動共有されることはありません。

1. **推奨:** 起動後の画面でAPIキーを入力する
2. **ローカル検証限定:** 自分のPC上だけでソースにAPIキーを一時的に埋め込む

公開GitHubリポジトリへ実APIキーをコミットしてはいけません。

---

## 0. まず確認すること

OrcaRouterのAPIキー一覧画面では、既存キーは次のように一部が伏せ字で表示されます。

```text
sk-orca-xxxx****xxxx
```

この伏せ字表示だけではAPI認証には使えません。

実行には、キー作成時に表示された**完全なAPIキー文字列**が必要です。

完全なキーを手元に保存していない場合は、OrcaRouter側で新しいAPIキーを作成してください。

このREADME、Issue、Chat、スクリーンショットなどへ完全なAPIキーを貼り付ける必要はありません。

スクリーンショットやログを共有するときは、APIキー欄を隠すか、次のようなダミー値へ置き換えてください。

```text
xxx-your-orcarouter-api-key-xxx
<API_KEY>
```

ローカルパスも個人のユーザー名を含む絶対パスではなく、次のように記載してください。

```text
<repository-root>
C:\Users\<USER>\...
/home/<USER>/...
```

---

# 1. Web版

対象ファイル:

```text
Web/app.js
```

ファイル冒頭付近に次の設定があります。

```javascript
const API_KEY_PLACEHOLDER = "xxx-your-orcarouter-api-key-xxx";

// LOCAL TEST ONLY:
const DEFAULT_API_KEY = "xxx-your-orcarouter-api-key-xxx";
```

## OrcaRouter Consoleからダウンロードしたキーのファイルを使う場合（推奨）

Web画面の **「キーのファイルを読込」** を押し、OrcaRouter Consoleからダウンロードしたキーのファイルを選択します。

ブラウザ側でファイル本文を読み、最初に見つかった完全な `sk-orca-...` 形式の文字列をAPI Key欄へ設定します。

対応しやすい形式:

- プレーンテキスト
- JSON
- .env形式
- その他、本文中に完全な `sk-orca-...` が含まれるテキストファイル

読み込み時にTraceへ出すのは、ファイル名・サイズ・マスク済みキーだけです。**ファイル本文と完全なAPIキーはTraceへ出力しません。**

この処理はブラウザ内で完結し、キーのファイル自体をGitHubやこのリポジトリへアップロードする必要はありません。

## 画面で直接入力する場合（推奨）

ソースは変更しません。

ブラウザでWebサンプルを開いたあと、API key欄へ実APIキーを入力します。

この方法なら、GitHub上のソースに実キーは残りません。

## ソースへ埋め込む場合（ローカル検証限定）

**変更するのは `DEFAULT_API_KEY` だけです。**

変更前:

```javascript
const DEFAULT_API_KEY = "xxx-your-orcarouter-api-key-xxx";
```

ローカルPC上でのみ、次のように変更します。

```javascript
const DEFAULT_API_KEY = "ここに自分の完全なAPIキー";
```

`API_KEY_PLACEHOLDER` は変更しないでください。

ブラウザを開くと、API key欄にその値が初期表示されます。

### 重要

Web版のJavaScriptへ実APIキーを埋め込むと、ブラウザの開発者ツールからキーを確認できます。

したがって、この方法は**localhostでの個人検証専用**です。

実APIキー入りのWeb版をGitHub Pagesや一般公開Webサーバーへ配置しないでください。

---

# 2. PowerShell版

対象ファイル:

```text
PowerShell/OrcaRouterChat.ps1
```

ファイル冒頭付近に次の設定があります。

```powershell
$script:ApiKeyPlaceholder = 'xxx-your-orcarouter-api-key-xxx'

# LOCAL TEST ONLY:
$script:DefaultApiKey = 'xxx-your-orcarouter-api-key-xxx'
```

## 画面で入力する場合（推奨）

ソースは変更しません。

PowerShellサンプルを起動すると、XAML画面上部にAPI Key欄があります。

そこへ実APIキーを入力してください。

## ソースへ埋め込む場合（ローカル検証限定）

**変更するのは `$script:DefaultApiKey` だけです。**

変更前:

```powershell
$script:DefaultApiKey = 'xxx-your-orcarouter-api-key-xxx'
```

ローカルPC上でのみ、次のように変更します。

```powershell
$script:DefaultApiKey = 'ここに自分の完全なAPIキー'
```

`$script:ApiKeyPlaceholder` は変更しないでください。

起動時にXAMLのAPI Key欄へ自動設定されます。

実行例:

```powershell
cd PowerShell
powershell.exe -STA -ExecutionPolicy Bypass -File .\OrcaRouterChat.ps1
```

---

# 3. Excel VBA版

対象ファイル:

```text
VBA/OrcaRouterSample.bas
```

モジュール冒頭付近に次の設定があります。

```vb
Private Const API_KEY_PLACEHOLDER As String = "xxx-your-orcarouter-api-key-xxx"

'LOCAL TEST ONLY:
Private Const DEFAULT_API_KEY As String = "xxx-your-orcarouter-api-key-xxx"
```

## Excelシートで入力する場合（推奨）

ソースは変更しません。

`SetupOrcaRouterSample` を実行すると、`OrcaRouter Chat` シートが作成されます。

APIキーの入力場所:

```text
B3セル
```

B3セルのダミー値を実APIキーへ置き換えて「送信」を押します。

## VBAソースへ埋め込む場合（ローカル検証限定）

**変更するのは `DEFAULT_API_KEY` だけです。**

変更前:

```vb
Private Const DEFAULT_API_KEY As String = "xxx-your-orcarouter-api-key-xxx"
```

ローカルPC上でのみ、次のように変更します。

```vb
Private Const DEFAULT_API_KEY As String = "ここに自分の完全なAPIキー"
```

`API_KEY_PLACEHOLDER` は変更しないでください。

その状態で `SetupOrcaRouterSample` を実行すると、B3セルへAPIキーが初期設定されます。

### 注意

VBAコードを含む `.xlsm` ファイルを第三者へ渡す場合、VBEから定数値を確認できるため、実APIキーを残さないでください。

---

# 4. GitHubへpushする前の確認

実APIキーをソースへ一時的に埋め込んだ場合は、push前に必ずダミー値へ戻してください。

3ファイルの設定がすべて次に戻っていることを確認します。

```text
xxx-your-orcarouter-api-key-xxx
```

Gitを使っている場合は、少なくとも次を確認します。

```bash
git status
git diff
```

OrcaRouterキーの一般的な先頭文字列も検索します。

```bash
git grep "sk-orca-"
```

何も表示されないことを確認してからcommit / pushしてください。

より広く確認する場合:

```bash
git grep -n -E "sk-orca-|API_KEY|DefaultApiKey|DEFAULT_API_KEY"
```

この検索ではダミー値や変数名も表示されるため、**完全な実APIキーが含まれていないこと**を目視確認します。

---

# 5. 誤ってGitHubへ実APIキーをpushした場合

最優先は、Git履歴の編集ではなく**APIキーの無効化・再発行**です。

1. OrcaRouterのAPIキー管理画面を開く
2. 誤って公開したAPIキーを無効化または削除
3. 新しいAPIキーを発行
4. ローカル設定を新しいキーへ差し替える
5. GitHub上のソースから旧キーを削除する

単に最新コミットから文字列を削除しても、過去のGit履歴に残っている可能性があります。

**一度公開されたキーは、漏えいしたものとして扱い、必ずローテーションしてください。**

---

# 6. どの方法を選ぶか

| 方法 | 手軽さ | GitHub漏えいリスク | 用途 |
|---|---:|---:|---|
| 画面で毎回入力 | 高い | 低い | 最初の動作確認 |
| ソースへ一時埋め込み | 高い | 高い | 自分のPCだけで繰り返し検証 |
| 環境変数 / Secret管理 | 中 | 低い | 継続利用・本番向け |

このリポジトリでは、学習時にコードを追いやすくするため「画面入力」と「ローカル限定の直書き」の両方を説明しています。

本番運用では、環境変数、Windows Credential Manager、CI/CD Secret、サーバー側Secret管理などへ移行してください。

# 7. 個人情報・ローカルパスのダミー化

README、Issue、Pull Request、Trace例、スクリーンショットへ次の情報をそのまま掲載しないでください。

- 完全なAPIキー
- 個人名
- メールアドレス
- 電話番号
- PCのユーザー名
- 個人PC固有の絶対ファイルパス
- 秘密情報を含むRaw JSONやTrace

例:

```text
NG:  C:\Users\actual-user-name\Documents\OrcaRouter-Samples
OK:  <repository-root>

NG:  /home/<USER>/OrcaRouter-Samples
OK:  <repository-root>

NG:  sk-orca-<real-secret-value>
OK:  xxx-your-orcarouter-api-key-xxx
```

このリポジトリのGitHub Actionsでは、現在のソースに実キーらしい文字列や代表的な個人環境パスが混入していないかを自動チェックします。ただし、自動チェックだけに依存せず、公開前に `git diff` とスクリーンショットを目視確認してください。

