# JavaScript 実行

ページ上で任意の JS を評価します。JavaScript 版の低レベル操作としての位置付けで、スクレイピングに関わる引数を一通り用意しています。

入口は次の 2 つです（いずれもタブ／コンテキスト単位）。

| プロトコル | メソッド | 内部プロトコル |
| --- | --- | --- |
| CDP | [`CDPContext.jsEval`](/api/cdp/CDPContext#jseval) | `Runtime.evaluate` / `Runtime.callFunctionOn` |
| BiDi | [`WebDriverBiDiContext.jsEval`](/api/bidi/WebDriverBiDiContext#jseval) | `script.evaluate` / `script.callFunction` |

::: tip
利用時は **名前付き引数**（`xxx:=yyy`）での呼び出しを推奨します。
:::

## 基本

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")

Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

' 例外時の代替値（IFERROR チック）
result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")
```

:::

`JavaScriptStr` は As Is で実行されます。エスケープの大半は `vbacollective-json` が担うため、手動エスケープは基本不要です。

## 3 つの実行パターン

`objectId` / `scriptHandle` や `objectArguments` の有無で、内部メソッドが切り替わります。

| パターン | CDP | BiDi | `JavaScriptStr` の書き方 |
| --- | --- | --- | --- |
| 1. 基準オブジェクトあり | `Runtime.callFunctionOn`（`objectId`） | `script.callFunction`（`scriptHandle`） | `function () { this... }` 形式 |
| 2. 引数だけ安全に渡す | `Runtime.callFunctionOn`（`objectArguments`） | `script.callFunction`（`objectArguments`） | 同上 |
| 3. ただのコード実行 | `Runtime.evaluate` | `script.evaluate` | 通常の式／文 |

```vb
' パターン1: 既存 objectId / handle に対して関数実行
result = t.jsEval( _
    "function () { return this.textContent; }", _
    objectId:=oid)          ' BiDi なら scriptHandle:=hid

' パターン3: 単純評価
result = t.jsEval("1 + 1")
```

`objectArguments` の組み立ては次を参照してください。

- CDP: [Runtime.CallArgument](https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#type-CallArgument)
- BiDi: [script.LocalValue](https://w3c.github.io/webdriver-bidi/#cddl-type-scriptlocalvalue)

## 主な引数

### 共通イメージ

| 目的 | CDP | BiDi |
| --- | --- | --- |
| 実行コード | `JavaScriptStr` | 同左 |
| 基準オブジェクト | `objectId` | `scriptHandle` |
| 関数引数 | `objectArguments` | 同左 |
| 例外時の代替値 | `IFEXCEPTION` | 同左 |
| iframe など実行場所 | `contextId`（`objectId` 指定時は無視） | `RealmTarget` |
| Promise 完了待ち | `awaitPromise` | 同左 |
| 人間操作としての偽装 | `userGesture` | `userActivation` |
| シリアライズ細調整 | `serializationOptions` | 同左 |
| 結果を待たない実行 | `RunAsyncCDP`（戻り値は実行 id） | `RunAsyncBiDi` |
| JS 例外で VBE 停止 | `StopException`（開発時向け） | 同左 |
| 通信エラーで停止 | `StopPipeError`（既定 `True`） | `StopBiDiError`（既定 `True`） |

### オブジェクト結果の受け取り方（CDP: `returnByValue` / BiDi: `Ownership`）

オブジェクト型の結果を「中身」で取るか、「ブラウザ内の参照 id」で取るかのスイッチです。**CDP と BiDi で True/False の意味が逆寄り**なので注意してください。

| | CDP `returnByValue` | BiDi `Ownership` |
| --- | --- | --- |
| 参照 id を得る（次回 `jsEval` に流用） | `False`（既定）→ `objectId` 文字列 | `True` → `handle` 文字列（`resultOwnership: "root"`） |
| 値として受け取る | `True` → `value` を試みる | `False`（既定）→ `value` を試みる（`resultOwnership: "none"`） |

- プリミティブ型の結果は、上記設定に関わらず中身が返ります
- 常に `objectId` / `handle` が欲しい場合は、JS 側の書き方を工夫してください
- 値受け取りで空オブジェクトになる場合は、対応していないデータ型のことがあります。CDP なら `returnByValue:=False`、BiDi なら `Ownership:=True` に切り替えてください
- `returnByValue:=True`（または BiDi で値受け取り）時、結果によって `Set` / `Let` が変わります。静的に決められないときは `Array` で受けて `VarType` 判定する手もあります

```vb
' CDP: オブジェクト参照を保持して次の jsEval に渡す
Dim oid As String
oid = t.jsEval("document.body", returnByValue:=False)
result = t.jsEval("function () { return this.tagName; }", objectId:=oid)

' BiDi: handle を保持
Dim hid As String
hid = t.jsEval("document.body", Ownership:=True)
result = t.jsEval("function () { return this.tagName; }", scriptHandle:=hid)
```

### その他のスイッチ

- **`awaitPromise`**: JS 結果が Promise でも完了まで待つ（`True`）
- **`userGesture` / `userActivation`**: 「人間が操作した」と偽装。スクレイピング対策向け
- **`allowUnsafeEvalBlockedByCSP`（CDP）**: CSP による外部 JS 規制を無視して実行許可
- **`serializationOptions`**: 意地でも `value` で受け取るための細かい調整。指定時は CDP の `returnByValue` は無視されます
- **`generatePreview`（CDP）**: `objectId` の中身をプレーンテキストでプレビュー。ログレベルを DEBUG 以下に。開発時向け
- **`RunAsyncCDP` / `RunAsyncBiDi`**: 結果を待たず実行 id のみ返す。「この操作、`alert` が発動するな…」という場面向け。回収は [低レイヤーガイド](/guides/extend-raw-protocol) / [イベント購読](/guides/events) 参照

## 例外とエラーの見分け

| 状況 | 確認方法 |
| --- | --- |
| JS 実行中の例外 | `IsError(result)`。詳細は `LastJavaScriptException`。面倒なら `IFEXCEPTION` |
| JS は成功したが結果がエラー値 | `IsError` で対処（`IFEXCEPTION` は「例外」時のみ機能） |
| 運用時の `StopException` | JS は例外を処理に使うことが多いため、**省略（止めない）を推奨**。止めるのは開発時向け |
| CDP / BiDi 通信自体のエラー | 既定で停止。第 3 引数系で制御 |

「この JS ならこういう返り値」という感覚がないと戻り値に困惑しやすいです。開発時はデバッグログで試行錯誤してください。

## スクリプトの追加・通知（CDP）

```vb
t.jsAddLib "https://cdn.example.com/lib.js"   ' URL
t.jsAddScript "C:\scripts\helper.js"          ' ローカルファイル
t.notify "処理が完了しました", 5               ' 表示秒数
```

## ダイアログ

アラート処理のデモは CDP / BiDi 双方にあります（`Demo_CDP.TestAlert` / `Demo_WebDriverBiDi.TestAlert`）。

BiDi ではセッション capability で自動 dismiss を無効化してから、イベント駆動で応答するパターンが典型です（[はじめに](/getting-started#sessioncapabilitiesrequest-とは) / [イベント購読](/guides/events)）。

CDP:

```vb
t.handleDialog True            ' Accept
t.handleDialog False, "入力"   ' prompt 用テキスト
```

## 関連

- [`CDPContext.jsEval`](/api/cdp/CDPContext#jseval) — [Runtime.evaluate](https://chromedevtools.github.io/devtools-protocol/tot/Runtime/#method-evaluate)
- [`WebDriverBiDiContext.jsEval`](/api/bidi/WebDriverBiDiContext#jseval) — [script.evaluate](https://w3c.github.io/webdriver-bidi/#command-script-evaluate)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [イベント購読](/guides/events)
