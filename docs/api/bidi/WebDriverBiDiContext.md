---
description: browsing context 単位の navigate・jsEval・ExecuteBiDi・CDP 変換など、BiDi のページ操作を解説します。
---

# WebDriverBiDiContext

BiDi の browsing context（タブ／ページ）です。Playwright の **Page** に相当します。

```vb
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext
t.navigate "https://example.com"
t.ThisWebDriverBiDiMode.CloseBrowser
```

親は [`ThisWebDriverBiDiMode`](#thiswebdriverbidimode)（[`WebDriverBiDiMode`](./WebDriverBiDiMode)）です。日常利用では設定シート経由の `StartBiDiModeContext` を推奨します。

要素のクリック／入力などは、`getElementByQuery` などで取得した [`WebDriverBiDiElement`](./WebDriverBiDiElement) で行います。CDP 独自の機能が必要な場合は、[`UpgradeBiDiPlus`](#upgradebidiplus) で同じタブを `CDPContext` に変換できます。

## 起動・再接続

### `OpenBrowserAndConnectTab`

```vb
Public Sub OpenBrowserAndConnectTab(caps As BrowserCapabilities)
```

ブラウザ起動・`session.new`・初回タブ（context）接続まで自動で行います。日常利用では設定シート経由で十分です。低レベルに起動したいときだけ直接呼んでください。

| 引数 | 意味 |
| --- | --- |
| `caps` | 起動時のオプションをまとめた [`BrowserCapabilities`](/api/cdp/BrowserCapabilities) |

ユーザーデータフォルダ名・ブラウザの種類・追加の起動引数・初期 URL などは、すべて `caps` で渡します。`session.new` に渡す `sessionCapabilitiesRequest` は、`caps.Set_BiDiSessionCapabilitiesRequest` で設定します。

```vb
BrowserCapabilities.Set_StartURL = "https://example.com"

Dim t As WebDriverBiDiContext
Set t = New WebDriverBiDiContext
t.OpenBrowserAndConnectTab BrowserCapabilities
```

Context 経由の起動は、常に「起動スプラッシュ画面」を経由してから実際の URL へ遷移します（`--app` 使用時のレースコンディション対策。詳細は [ページ遷移](/guides/navigation)）。

`sessionCapabilitiesRequest` の詳細は [はじめに](/getting-started#sessioncapabilitiesrequest-とは)。

### `reattach`

```vb
Public Function reattach(reattachInfo As ReattachOptions) As Boolean
```

Excel テーブルに記録されたメイン BiDi context へ再接続を試みます。通信経路（Pipe / Port）は、[`ReattachOptions`](/api/cdp/ReattachOptions) で指定します。

| 引数 | 意味 |
| --- | --- |
| `reattachInfo` | 再接続時のオプションをまとめた [`ReattachOptions`](/api/cdp/ReattachOptions)。制御経路の指定が必須。新しい BiDi-CDP Mapper が起動されたときに使う `Set_BiDiSessionCapabilitiesRequest` なども、ここで設定します |

**戻り値:** 既存 context への再接続成功可否。

```vb
ReattachOptions.Set_PipeMode = ShSetting01_StartBrowser.CurrentUserName

Dim t As New WebDriverBiDiContext
If Not t.reattach(ReattachOptions) Then Exit Sub
```

::: tip 注意
- パイプが生きていない場合は、このメソッドから再開できません。記録中の context が既に閉じている場合は `False` になります
- `ReattachOptions` は使い切りです。`reattach` を呼ぶたびに、制御経路を指定し直してください
:::

詳細は [再接続](/guides/reattach) / [WebSocket モード](/websocket/capabilities)。

## ナビ・待機

### `navigate`

```vb
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
```

URL を開き、指定の読み込み条件まで待ちます。内部では `browsingContext.navigate` のあと [`WaitEvents`](#waitevents) を呼びます（`till:=Nowait` なら待ちません）。

| 引数 | 意味 |
| --- | --- |
| `strURL` | 遷移先 URL |
| `till` | 待機条件。既定は `isComplete`。`Nowait`（v3.2.0〜）で待たずに戻る |

| `ReadyState` | BiDi の `wait` |
| --- | --- |
| `isLoading` | `"none"` |
| `isInteractive` | `"interactive"` |
| `isComplete` | `"complete"` |

```vb
t.navigate "https://example.com"                          ' 完了まで待つ（既定）
t.navigate "https://example.com/heavy", isInteractive     ' interactive で先に進む
t.navigate "https://example.com", Nowait                  ' 待たずに次へ進む（v3.2.0〜）
```

::: tip CDP側との違い
CDP の [`CDPContext.navigate`](/api/cdp/CDPContext#navigate) は v3.2.0 で `referrer` / `transitionType` 等の引数が追加されましたが、BiDi の `browsingContext.navigate` にはそれに相当する概念がないため、こちらの `navigate` のシグネチャは変わっていません。
:::

詳細は [ページ遷移](/guides/navigation)。

### `Wait`

```vb
Public Sub Wait(Optional till As ReadyState = isComplete)
```

現在ページの `document.readyState` を**ポーリング**して、指定状態になるまで待ちます。`navigate` 後だけでなく、クリック後の再読み込み待ちなどにも使えます。新しいタブなど、いつ遷移が起きたか把握できない場面ではこちらが安全です。

| 引数 | 意味 |
| --- | --- |
| `till` | 待機する `document.readyState`。既定は `isComplete` |

```vb
t.Wait                              ' 完了待ち（既定）
t.Wait isInteractive
```

::: tip
- `till:=isInteractive` のとき、すでに `complete` まで進んでいればそのまま成功扱いで抜けます
- 一般的な読み込みステータスのみ対応です。SPA などの特殊な待機は別途実装が必要です
:::

### `WaitEvents`（v3.2.0〜）

```vb
Public Function WaitEvents(Optional WaitMode As ReadyState = isComplete, Optional WaitError As Boolean = True) As Boolean
```

`Wait` のポーリング版に対し、こちらは BiDi イベント（`browsingContext.navigationStarted` / `domContentLoaded` / `load`）を**受動的に待つ**版です。`navigate` が内部で使っているのはこちらです。CDP側の [`CDPContext.WaitEvents`](/api/cdp/CDPContext#waitevents) と同じ設計・同じ注意点（`ResetWaitState` を先に呼ぶ必要がある）です。

| 引数 | 意味 |
| --- | --- |
| `WaitMode` | 待機する `ReadyState`。既定は `isComplete` |
| `WaitError` | タイムアウト時に例外を出すか。既定は `True` |

```vb
t.ResetWaitState
t.getElementByQuery("a.next").click   ' または UpgradeBiDiPlus 経由のクリック
t.WaitEvents WaitMode:=isComplete
```

### `ResetWaitState`（v3.2.0〜）

```vb
Public Sub ResetWaitState()
```

[`WaitEvents`](#waitevents) 用の内部状態をリセットします。`navigate` を経由しない遷移トリガーの直前に呼んでください。

### `ReadyState`

[`CDPContext` と同じ列挙](/api/cdp/CDPContext#readystate)です（`Nowait`（v3.2.0〜） / `isLoading` / `isInteractive` / `isComplete`）。

## JavaScript

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, _
    Optional scriptSharedId As String, _
    Optional objectArguments As Variant, _
    Optional IFEXCEPTION As Variant, _
    Optional RealmTarget As String, _
    Optional Ownership As Boolean, _
    Optional awaitPromise As Boolean, _
    Optional userActivation As Boolean, _
    Optional serializationOptions As Dictionary, _
    Optional returnByDetails As Boolean, _
    Optional RunAsyncBiDi As Boolean, _
    Optional StopException As Boolean, _
    Optional StopBiDiError As Boolean = True) As Variant
```

ページ上で JavaScript を評価します（`script.evaluate` / `script.callFunction`）。利用時は **名前付き引数**を推奨します。

| 引数 | 意味 |
| --- | --- |
| `JavaScriptStr` | 実行する JS（そのまま送る） |
| `scriptSharedId` | `this` にする要素の `script.SharedId`。指定時は `script.callFunction`（`function(){ this... }` 形式）。**要素専用**です。要素以外のオブジェクトを使う場合は、`objectArguments` で `script.Handle` を渡します |
| `objectArguments` | 引数（`Collection` / `Array(...)` / 固定長 `Dictionary` 型 1 次元配列）。指定時も `callFunction` 優先。固定長配列の方がパフォーマンスが良い |
| `IFEXCEPTION` | JS 例外時の代替値（IFERROR チック） |
| `RealmTarget` | iframe 等の realm を指定して実行したいとき |
| `Ownership` | `True` で `resultOwnership: "root"`。要素以外のオブジェクトは `script.Handle` で返り、次回の `objectArguments` に流用できます |
| `awaitPromise` | Promise 完了まで待つか |
| `userActivation` | 人間操作として偽装するか（スクレイピング対策向け） |
| `serializationOptions` | 意地でも `value` で受け取るための細かい調整 |
| `returnByDetails` | `True` で、値の代わりに [`JsEvalResult`](/api/cdp/JsEvalResult)（型・`sharedId`・`handle`・例外情報など）を返します。`Set` で受けてください |
| `RunAsyncBiDi` | `True` で結果を待たず実行 id のみ返す |
| `StopException` | JS 例外で停止するか（開発時向け） |
| `StopBiDiError` | BiDi 通信エラーで停止するか。既定は `True` |

```vb
Dim result As Variant
result = t.jsEval("document.title")
Debug.Print result

result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")

' 要素（node）は script.SharedId で返る。this にして呼び出す
Dim sid As String
sid = t.jsEval("document.body")
result = t.jsEval("function () { return this.tagName; }", scriptSharedId:=sid)

' 要素以外（配列など）は Ownership:=True で script.Handle を受け取り、引数で渡す
Dim hid As String, args(0) As Dictionary
hid = t.jsEval("[1, 2, 3]", Ownership:=True)
Set args(0) = New Dictionary: args(0).Add "handle", hid
result = t.jsEval("function (a) { return a.length; }", objectArguments:=args)

' 値・型・参照 ID などをまとめて受け取る
Dim r As JsEvalResult
Set r = t.jsEval("document.body", returnByDetails:=True)
Debug.Print r.ResultType, r.sharedId          ' node / (sharedId)
```

例外時は `IsError(result)` で判定します。詳細は [JavaScript 実行](/guides/javascript)。

## 要素の取得

検索結果を [`WebDriverBiDiElement`](./WebDriverBiDiElement) で返します。クリック・入力・存在待ちなど、`CDPElement` とほぼ同じメソッドで操作できます。

```vb
Public Function getElementByID(strID As String) As WebDriverBiDiElement
Public Function getElementByQuery(strQuery As String) As WebDriverBiDiElement
Public Function getElementsByQuery(strQuery As String) As Collection
Public Function getElementByXPath(strXPath As String) As WebDriverBiDiElement
Public Function getElementsByXPath(strXPath As String) As Collection
```

```vb
t.getElementByQuery("input[name='q']").value = "検索語"
t.getElementByXPath("//button[text()='検索']").click

Dim item As WebDriverBiDiElement
For Each item In t.getElementsByQuery("ul.result > li")
    Debug.Print item.innerText
Next
```

- 複数系（`getElementsBy*`）の `Collection` は1始まりです。検索自体に失敗した場合は `Nothing` を返します
- iframe 内の要素は、[`WebDriverBiDiMode.getTab`](./WebDriverBiDiMode) の `maxDepth:=1` で iframe を取得してから検索します
- Shadow DOM は、ホスト要素の `GetShadowRoot` で1段ずつ辿ります

詳しくは [`WebDriverBiDiElement`](./WebDriverBiDiElement) / [要素の取得](/guides/selectors) を参照してください。

## BiDi+ への橋渡し

### `UpgradeBiDiPlus`

```vb
Public Function UpgradeBiDiPlus() As CDPContext
```

同じタブを [`CDPContext`](/api/cdp/CDPContext) に変換します。`CDPElement` によるクリック／入力などが使えます。失敗時は `Nothing` です。

`browsingContext.getTree` 等で得られる `context` が CDP の `targetId` と一致する仕様を利用した変換術で（内部で `Target.attachToTarget` により新規 `sessionId` を発行）、Chromium 独自拡張 `BiDi+`（[`chromium-bidi`](https://github.com/GoogleChromeLabs/chromium-bidi#bidi) 由来の命名）が担う3コマンドは、実質この変換だけで賄えます。

| BiDi+ | 相当する処理 |
| --- | --- |
| `goog:cdp.sendCommand` | `CDPContext.ExecuteCDP` / `ExecuteCDPAsync` |
| `goog:cdp.getSession` | `Target.attachToTarget`（本メソッド内部で実行） |
| `goog:cdp.resolveRealm` | `Runtime.executionContextCreated` / `Destroyed` / `Cleared` |

```vb
Dim cdp As CDPContext
Set cdp = t.UpgradeBiDiPlus
cdp.getElementByQuery("button").click
```

関連: [要素の取得](/guides/selectors) / [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)

## イベント購読（コンテキスト単位、v3.1.1.1〜）

### `SubscribeBiDiEvent`

```vb
Property Set SubscribeBiDiEvent(Optional subscribe As Boolean = True, events As Collection)
Property Let SubscribeBiDiEvent(Optional subscribe As Boolean = True, events)
```

`session.subscribe` / `session.unsubscribe` を、**このタブ（context）単位**で実行します。`ExecuteBiDi` と同じく params に `context` が自動付与されるため、他のタブの購読状態には影響しません。

| 引数 | 意味 |
| --- | --- |
| `subscribe` | `True`（既定）で購読、`False` で購読解除 |
| `events` | SET: イベント名の `Collection`／LET: 1次元配列または文字列1件 |

```vb
' このタブだけネットワーク系イベントを購読する例（他のタブは購読しないまま操作を続けられる）
t.SubscribeBiDiEvent = Array("network.beforeRequestSent", "network.responseCompleted")
```

::: tip ブラウザ全体の購読との使い分け
全タブ共通で購読したい場合は、これまで通り [`ThisWebDriverBiDiMode.SubscribeBiDiEvent`](./WebDriverBiDiMode#subscribebidievent) を使ってください。パフォーマンスを気にするなら、必要なタブだけに絞れるこちらの Context 単位版がおすすめです。
:::

詳細は [イベント購読](/guides/events)。

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

`ExecuteBiDiAsync` の結果回収は親の [`TakeResultBiDi`](./WebDriverBiDiMode#takeresultbidi) で行います。

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
    t.ThisWebDriverBiDiMode.TakeEvents
    Dim raw As String
    raw = t.ThisWebDriverBiDiMode.TakeResultBiDi(cmdId)
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

### `ThisWebDriverBiDiMode`

```vb
Property Get ThisWebDriverBiDiMode() As WebDriverBiDiMode
```

`WebDriverBiDiContext` 上で、セッション／ブラウザ単位の制御もしたいときに使います。親の [`WebDriverBiDiMode`](./WebDriverBiDiMode) への参照です。

タブ操作（`navigate` / `jsEval` など）は Context 側、タブ一覧／終了／イベント購読（`CloseBrowser` / `newTab` / `BiDiEvents` など）はこちら経由、という使い分けになります。

```vb
t.ThisWebDriverBiDiMode.newTab
t.ThisWebDriverBiDiMode.TimeOutSecond = 60
t.ThisWebDriverBiDiMode.CloseBrowser
```

::: tip 拡張クラスでイベントを受け取るには
このクラスにはイベントがありません。自作の拡張クラスでは、親の [`WebDriverBiDiMode`](./WebDriverBiDiMode) の `ExEvent` を `WithEvents` で受け取ってください。
:::

## タイムアウト

`TimeOutSecond` はこのクラス自体にはなく、親の [`WebDriverBiDiMode.TimeOutSecond`](./WebDriverBiDiMode#timeoutsecond) で設定します。

```vb
t.ThisWebDriverBiDiMode.TimeOutSecond = 60
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
- [`CDPContext`](/api/cdp/CDPContext)（`UpgradeBiDiPlus` 経由）
- [CDP と BiDi](/concepts/cdp-vs-bidi)
- [ページ遷移](/guides/navigation)
- [JavaScript 実行](/guides/javascript)
- [要素の取得](/guides/selectors)
- [イベント購読](/guides/events)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [タイムアウト設定方法について](/guides/timeout)
- [再接続](/guides/reattach)
