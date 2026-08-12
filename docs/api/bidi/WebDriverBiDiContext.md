---
description: browsing context 単位の navigate・jsEval・ExecuteBiDi・CDP 変換など、BiDi のページ操作を解説します。
---

# WebDriverBiDiContext

BiDi の browsing context（タブ／ページ）です。Playwright の **Page** に相当します。

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext
t.navigate "https://example.com"
t.InheritanceWebDriverBiDiMode.quit
```

親は [`InheritanceWebDriverBiDiMode`](#inheritancewebdriverbidimode)（[`WebDriverBiDiMode`](./WebDriverBiDiMode)）です。日常利用では設定シート経由の `StartBiDiModeContext` を推奨します。

要素のクリック／入力など高レベル API は CDP 側が充実しています。必要なら [`ConvertToCDPContext`](#converttocdpcontext) で変換してください。

## 起動・再接続

### `StartBiDiModeAndConnectTab`

```vb
Public Sub StartBiDiModeAndConnectTab( _
    Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional userProfile As String, _
    Optional addArgs As String, _
    Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary _
)
```

ブラウザ起動・`session.new`・初回タブ（context）接続まで自動で行います。日常利用では設定シート経由で十分です。低レベルに起動したいときだけ直接呼んでください。

| 引数 | 意味 |
| --- | --- |
| `Name` | ブラウザ名（現時点では Chrome / Edge） |
| `appUrl` | 起動時に開く URL |
| `userProfile` | `--user-data-dir` 用のユーザーディレクトリ名 |
| `addArgs` | 追加の起動引数 |
| `KioskMode` | Edge キオスクモード |
| `sessionCapabilitiesRequest` | `session.new` の params。事前に `Dictionary` で組み立てる |

```vb
Dim t As New WebDriverBiDiContext
t.StartBiDiModeAndConnectTab "chrome", userProfile:="MyUser"
```

`sessionCapabilitiesRequest` の詳細は [はじめに](/getting-started#sessioncapabilitiesrequest-とは)。

### `reattach`

```vb
Public Function reattach( _
    userProfile As String, _
    Optional sessionCapabilitiesRequest As Dictionary, _
    Optional WebSocketMode As CDPCoreViaWebSocket _
) As Boolean
```

Excel テーブルに記録されたメイン BiDi context へ再接続を試みます。

| 引数 | 意味 |
| --- | --- |
| `userProfile` | 再アタッチしたいユーザー名（`user-data-dir` に基づく） |
| `sessionCapabilitiesRequest` | 新しい BiDi-CDP Mapper が起動されたときだけ `session.new` に適用 |
| `WebSocketMode` | WebSocket で BiDi 制御する場合、接続済みの `CDPCoreViaWebSocket` を指定 |

**戻り値:** 既存 context への再接続成功可否。

```vb
Dim t As New WebDriverBiDiContext
If Not t.reattach(ShSetting01_StartBrowser.CurrentUserName) Then Exit Sub
```

::: tip 注意
パイプが生きていない場合は、このメソッドから再開できません。記録中の context が既に閉じている場合も `False` になります。
:::

詳細は [再接続](/guides/reattach) / [WebSocket モード](/websocket/capabilities)。

## ナビ・待機

### `navigate`

```vb
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
```

URL を開き、指定の読み込み条件まで待ちます。内部では `browsingContext.navigate` の `wait` にマッピングします。

| 引数 | 意味 |
| --- | --- |
| `strURL` | 遷移先 URL |
| `till` | 待機条件。既定は `isComplete` |

| `ReadyState` | BiDi の `wait` |
| --- | --- |
| `isLoading` | `"none"` |
| `isInteractive` | `"interactive"` |
| `isComplete` | `"complete"` |

```vb
t.navigate "https://example.com"                          ' 完了まで待つ（既定）
t.navigate "https://example.com/heavy", isInteractive     ' interactive で先に進む
```

詳細は [ページ遷移](/guides/navigation)。

### `wait`

```vb
Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
```

現在ページの `document.readyState` が指定状態になるまで待ちます。`navigate` 後だけでなく、クリック後の再読み込み待ちなどにも使えます。

| 引数 | 意味 |
| --- | --- |
| `till` | 待機する `document.readyState`。既定は `isComplete` |
| `dbgState` | `True` で待機中の ReadyState をログへ出し、`jsEval` の例外も止めずに継続 |

```vb
t.wait                              ' 完了待ち（既定）
t.wait isInteractive
t.wait isComplete, dbgState:=True
```

::: tip
- `till:=isInteractive` のとき、すでに `complete` まで進んでいればそのまま成功扱いで抜けます
- 一般的な読み込みステータスのみ対応です。SPA などの特殊な待機は別途実装が必要です
:::

### `ReadyState`

[`CDPContext` と同じ列挙](/api/cdp/CDPContext#readystate)です（`isLoading` / `isInteractive` / `isComplete`）。

## JavaScript

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, _
    Optional scriptHandle As String, _
    Optional objectArguments As Variant, _
    Optional IFEXCEPTION As Variant, _
    Optional RealmTarget As String, _
    Optional Ownership As Boolean, _
    Optional awaitPromise As Boolean, _
    Optional userActivation As Boolean, _
    Optional serializationOptions As Dictionary, _
    Optional RunAsyncBiDi As Boolean, _
    Optional StopException As Boolean, _
    Optional StopBiDiError As Boolean = True) As Variant
```

ページ上で JavaScript を評価します（`script.evaluate` / `script.callFunction`）。利用時は **名前付き引数**を推奨します。

| 引数 | 意味 |
| --- | --- |
| `JavaScriptStr` | 実行する JS（そのまま送る） |
| `scriptHandle` | 基準ハンドル。指定時は `script.callFunction`（`function(){ this... }` 形式） |
| `objectArguments` | 引数（`Collection` / `Array(...)` / 固定長 `Dictionary` 型 1 次元配列）。指定時も `callFunction` 優先。固定長配列の方がパフォーマンスが良い |
| `IFEXCEPTION` | JS 例外時の代替値（IFERROR チック） |
| `RealmTarget` | iframe 等の realm を指定して実行したいとき |
| `Ownership` | `True` で `resultOwnership: "root"`（次回の `scriptHandle` に流用可能な handle を返す） |
| `awaitPromise` | Promise 完了まで待つか |
| `userActivation` | 人間操作として偽装するか（スクレイピング対策向け） |
| `serializationOptions` | 意地でも `value` で受け取るための細かい調整 |
| `RunAsyncBiDi` | `True` で結果を待たず実行 id のみ返す |
| `StopException` | JS 例外で停止するか（開発時向け） |
| `StopBiDiError` | BiDi 通信エラーで停止するか。既定は `True` |

```vb
Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")

' handle を保持して次の呼び出しに流用
Dim hid As String
hid = t.jsEval("document.body", Ownership:=True)
result = t.jsEval("function () { return this.tagName; }", scriptHandle:=hid)
```

例外時は `IsError(result)` で判定します。詳細は [JavaScript 実行](/guides/javascript)。

## CDP への橋渡し

### `ConvertToCDPContext`

```vb
Public Function ConvertToCDPContext() As CDPContext
```

同じタブを [`CDPContext`](/api/cdp/CDPContext) として操作できます。`CDPElement` によるクリック／入力などが使えます。失敗時は `Nothing` です。

```vb
Dim cdp As CDPContext
Set cdp = t.ConvertToCDPContext
cdp.getElementByQuery("button").click
```

関連: [要素の取得](/guides/selectors)

## プロトコル

### `ExecuteBiDi` / `ExecuteBiDiAsync`

```vb
Public Function ExecuteBiDi(methodName As String, _
    Optional params As Dictionary, _
    Optional StopBiDiError As Boolean = True) As BiDiCDPJson

Public Function ExecuteBiDiAsync(methodName As String, _
    Optional params As Dictionary, _
    Optional StopError As Boolean = True) As Long
```

この context 向け BiDi コマンドです。params に **`context` を自動付与**してから親 Mode へ渡します。BiDi+（`goog:cdp.*`）もここから呼べます。

| 引数 | 意味 |
| --- | --- |
| `methodName` | メソッド名（例: `"browsingContext.navigate"`） |
| `params` | params の `Dictionary`。省略時は空 `{}` に `context` だけ付く |
| `StopBiDiError` / `StopError` | 失敗時に停止するか。既定は `True` |

`ExecuteBiDiAsync` の結果回収・蓄積上限は親の [`TakeResultBiDi`](./WebDriverBiDiMode#takeresultbidi) / [`SetLimitBiDi`](./WebDriverBiDiMode#setlimitbidi) で行います。

```vb
Dim params As New Dictionary
params.Add "url", "https://example.com"
params.Add "wait", "complete"

Dim result As BiDiCDPJson
Set result = t.ExecuteBiDi("browsingContext.navigate", params)

' 非同期
Dim cmdId As Long
cmdId = t.ExecuteBiDiAsync("browsingContext.navigate", params)

Do
    t.InheritanceWebDriverBiDiMode.TakeEvents
    Dim raw As String
    raw = t.InheritanceWebDriverBiDiMode.TakeResultBiDi(cmdId)
    If LenB(raw) Then Exit Do
    DoEvents
Loop
```

詳細は [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)。

## タブ情報・生存

### `context`

```vb
Property Get context() As String
Property Let context(param As String)
```

このタブの browsing context id です。`browsingContext.getTree` などで得られる値と対応します。日常の読み取りは Get、Mode 側から渡すときは Let が使われます。

```vb
Debug.Print t.context
```

[再接続](/guides/reattach) や `serializeMainTab` で Excel に残る値と対応します。

### `isContextClosed`

```vb
Property Get isContextClosed() As Boolean
```

この context が閉じられたかです。`browsingContext.contextDestroyed` を受け取ると `True` になります。`True` ならこのオブジェクトはもう使えないので破棄してください。

```vb
If t.isContextClosed Then
    ' オブジェクトを捨てて、必要なら getTab / newTab で取り直す
End If
```

## 親セッション

### `InheritanceWebDriverBiDiMode`

```vb
Property Get InheritanceWebDriverBiDiMode() As WebDriverBiDiMode
```

`WebDriverBiDiContext` 上で、セッション／ブラウザ単位の制御もしたいときに使います。親の [`WebDriverBiDiMode`](./WebDriverBiDiMode) への参照です。

タブ操作（`navigate` / `jsEval` など）は Context 側、タブ一覧／終了／イベント購読（`quit` / `newTab` / `BiDiEvents` など）はこちら経由、という使い分けになります。

```vb
t.InheritanceWebDriverBiDiMode.newTab
t.InheritanceWebDriverBiDiMode.TimeOutSecond = 60
t.InheritanceWebDriverBiDiMode.quit
```

### `InheritanceWebDriverBiDiCore`

```vb
Property Get InheritanceWebDriverBiDiCore() As WebDriverBiDiCore
```

内部の `WebDriverBiDiCore` への参照です。通常は Mode／Context の公開 API で十分です。

## タイムアウト

`TimeOutSecond` はこのクラス自体にはなく、親の [`WebDriverBiDiMode.TimeOutSecond`](./WebDriverBiDiMode#timeoutsecond) で設定します。

```vb
t.InheritanceWebDriverBiDiMode.TimeOutSecond = 60
```

詳細は [タイムアウト設定方法について](/guides/timeout)。

## デバッグ

### `printMsg`

```vb
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, _
    Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
```

デバッグ／ログ出力です。通常はフレームワーク内部から呼ばれます。

## 関連

- [`WebDriverBiDiMode`](./WebDriverBiDiMode)
- [`CDPContext`](/api/cdp/CDPContext)（`ConvertToCDPContext` 経由）
- [CDP と BiDi](/concepts/cdp-vs-bidi)
- [ページ遷移](/guides/navigation)
- [JavaScript 実行](/guides/javascript)
- [要素の取得](/guides/selectors)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [タイムアウト設定方法について](/guides/timeout)
- [再接続](/guides/reattach)
