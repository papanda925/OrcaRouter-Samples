# UI behavior contract / request-state rules

この文書は Web / PowerShell-WPF / Excel VBA の3実装で共通に守るUI動作を定義します。

2026-09-02 のPowerShell実機確認で、送信後に質問が回答欄へ一瞬表示された後に消え、
エラー時にTraceへ自動遷移し、回答とDeveloper Informationが確認できない問題が見つかりました。
同じ設計パターンがWeb/VBAにも一部存在していたため、個別修正ではなく3実装共通の契約として固定します。

## 1. 原因

今回の問題はAPI通信そのものより、**requestの状態と画面表示の状態を分離していなかったこと**が主因です。

以前の実装では次の処理が混在していました。

1. Send直後に「送信中の質問」を結果欄へ仮表示する
2. 成功した会話履歴だけを別の配列/変数へ保持する
3. エラーになると「確定済み履歴」から結果欄を再描画する
4. エラー診断のためTraceへ自動遷移する
5. Developer Informationは主に成功完了時に更新する

この組合せにより、Send直後に見えた質問は「確定済み履歴」ではないため、
エラー時の再描画で消えました。さらにTraceへの自動遷移により、
利用者には「回答が消えて処理が止まった」ように見えました。

Prompt example / Mode変更にも同じ種類の問題がありました。
**選択操作だけでQuestionを書き換える副作用**があると、利用者が編集中の質問を失う可能性があります。

## 2. 共通状態

3実装はRequestの状態を同じ意味で扱います。**conversation historyはAPIへ渡す内部状態であり、回答欄へそのまま連結表示する必要はありません。**

PowerShell版では、長文時の可読性と応答性を優先し、「回答」には最新のAssistant出力だけを表示します。Request / Response等の診断情報はDeveloper / Traceへ分離します。

| State | 内部Conversation state | 主表示 | Status |
|---|---|---|---|
| Ready | 成功済み履歴を保持 | 前回結果または初期表示 | Ready |
| Sending | まだ今回turnをcommitしない | 未確定Questionを回答として表示しない | 回答待ち |
| Streaming | まだ今回turnをcommitしない | 受信済みpartial Assistant | 受信中 |
| Success | Question + Assistantを履歴へcommit | 今回のAssistant結果 | Completed |
| Error | 失敗turnはcommitしない | ERROR内容 | Error |
| New Chat | 履歴を空にする | 初期表示 | Ready/New chat |

## 3. 必須ルール

### 3.1 Send直後に未確定の質問を回答欄へ表示しない

SendしたことはStatusで示します。PowerShell版では送信時に前回Answerをクリアし、「送信済み・回答待ち...」を表示します。

Streamingは実際にAssistantのdeltaを受信した時点からpartial answerを表示します。質問文そのものを回答欄へ繰り返し表示しません。

### 3.2 失敗したturnは履歴へcommitしない

APIエラー、timeout、JSON parse error、入力検証エラーは、
次回APIへ送るconversation historyへ追加しません。

ただし利用者が原因を確認できるよう、ERROR内容は主結果領域へ表示します。質問文自体はQuestion欄に残っているため、PowerShell版のAnswerへ重複表示しません。

### 3.3 Errorでも結果画面を消さない

エラー発生時に主結果領域を空のままにしません。PowerShell版では回答欄にERROR内容だけを表示します。

表示例:

```text
ERROR: HTTP 429 ...
```

Questionが空で入力検証に失敗した場合も、ERROR自体は必ず表示します。

### 3.4 ErrorでTraceへ自動遷移しない

結果の主表示はAnswer / Conversationです。
エラー時も利用者を別タブ・別領域へ強制移動しません。

Developer / Traceは利用者が必要に応じて開く診断領域です。

### 3.5 Developer Informationは成功・失敗の両方で更新する

可能な範囲で次を残します。

- HTTP Status
- Elapsed
- Model
- Prompt Tokens
- Completion Tokens
- Total Tokens
- Cost
- Request JSON
- Response JSON / Error body

HTTPレスポンス前に失敗した場合は、存在しない値を推測せず
`-` / `(not available)` / `(not returned)` とします。

### 3.6 Prompt exampleは明示的に適用する

Prompt example / Prompt templateは選択しただけではQuestionを書き換えません。

```text
例を選択
  ↓
「質問欄に挿入 / Insert prompt」
  ↓
Questionへ反映
  ↓
利用者が編集
  ↓
Send
```

### 3.7 Mode変更でQuestionを書き換えない

Chat / Streaming / Tool CallingのMode選択は通信方式の選択です。
利用者が入力中のQuestionを破棄してはいけません。

Mode固有の例文が必要な場合はPrompt example等の明示操作で提供します。

### 3.8 New Chatが消すのはconversation history

New ChatはAPIへreset電文を送りません。

原則として次は保持します。

- API Key
- Model
- Mode
- Developer Information / Trace（診断のため残してよい）

conversation historyだけを空にします。

## 4. 履歴

履歴はアプリ側が保持する**内部Conversation state**で、最大10往復です。PowerShell版では履歴全体を回答欄へ連結表示しません。

```text
turn = user + assistant
```

11往復目が成功したら、最古の成功turnを削除します。

Web / PowerShell / VBAのChat、Streaming、Tool Callingは、
可能な範囲で同じ成功済み履歴を次回Requestの `messages` へ含めます。

## 4.1 PowerShell Runspace境界のCollectionを正規化する

PowerShell版では、UI RunspaceからBackground Runspaceへ渡した値が、空・単一要素・配列で形を変える可能性を考慮します。

特に `History` 件数は `@($History).Count` のような直接参照に頼らず、nullを除外して数える共通Helperを使います。

Workerで例外が発生した場合は、MessageだけでなくScript line / Position / Script stackもDeveloper / Traceへ残し、HTTP処理とローカル後処理のどちらで失敗したか判別できるようにします。

## 5. 再発防止

レビューでは「コントロールが存在するか」「関数名があるか」だけでなく、
**状態遷移の結果として利用者に何が見えるか**を確認します。

必須回帰シナリオ:

1. Questionを入力してSend → 応答待ち中にQuestionが結果欄へ点滅しない
2. 成功 → 最新のAssistant回答が表示され、Questionを回答欄へ重複表示しない
3. APIエラー → ERROR内容が結果欄へ残り、失敗turnを履歴へ追加しない
4. Error → Traceへ勝手に遷移しない
5. Error → Developerに取得可能な診断情報が残る
6. Prompt exampleを選択 → Questionは変わらない
7. Apply/Insert → 初めてQuestionが変わる
8. Mode変更 → Questionは変わらない
9. New Chat → 内部履歴だけ消える
10. 11回成功 → 内部履歴は最大10往復
11. 長文Question / Answer → 各TextBox内部でスクロールし、本文量だけでフォーム全体を無制限に伸ばさない
12. 画面の高さが足りない → 右側のページスクロールでRESULT / Statusまで到達できる
13. 回答待ち → Busy表示が見え、WPF UIは操作可能なまま

PowerShellはWPFの実UI自己テストでも確認します。
Web/VBAはCIのUI contract検査と各実装の自己テスト可能部分で確認します。

## 6. レビュー観点

今後の変更では最低限、次の4観点を分けて確認します。

- Beginner UX: 初見利用者に操作の意味が分かるか
- API / Security: Request、秘密情報、エラー情報の扱いが正しいか
- Reliability: 成功/失敗/timeout/Streamingの状態遷移が壊れないか
- Developer Experience: Request/Response/Token/Cost/Traceが学習に役立つか

静的な「存在確認」がすべてPASSしても、状態遷移レビューがPASSしなければ完了とはしません。
