# WebDriverBiDiMode

BiDi セッション／ブラウザ相当です。Playwright の **Browser** に対応します。

```vb
Dim mode As WebDriverBiDiMode
Set mode = ShSetting01_StartBrowser.StartBiDiMode
Dim tab As WebDriverBiDiContext
Set tab = mode.getTab(setMain:=True)
tab.navigate "https://example.com"
mode.quit
```

タブからすぐ始めたい場合は設定シートの `StartBiDiModeContext` → [`WebDriverBiDiContext`](./WebDriverBiDiContext) が簡単です。

## 起動・再接続・終了

### `StartBiDiMode`

```vb
Public Sub StartBiDiMode( _
    Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional userProfile As String, _
    Optional addArgs As String, _
    Optional KioskMode As edgeKioskType, _
    Optional sessionCapabilitiesRequest As Dictionary _
)
```

ブラウザを WebDriver BiDi として起動し、`session.new` まで行います。日常利用では設定シート経由を推奨します。

| 引数 | 意味 |
| --- | --- |
| `Name` | ブラウザ名（現時点では Chrome / Edge） |
| `appUrl` | 起動時に開く URL（`--app` 相当） |
| `userProfile` | `--user-data-dir` 用のユーザーディレクトリ名 |
| `addArgs` | 追加の起動引数 |
| `KioskMode` | Edge キオスクモード |
| `sessionCapabilitiesRequest` | `session.new` の params。事前に `Dictionary` で組み立てる |

```vb
Dim caps As New Dictionary
' ... capabilities を組み立て ...
Dim mode As New WebDriverBiDiMode
mode.StartBiDiMode "chrome", userProfile:="MyUser", sessionCapabilitiesRequest:=caps
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

既存の BiDi 接続へ再接続を試みます。

| 引数 | 意味 |
| --- | --- |
| `userProfile` | 再アタッチしたいユーザー名（`user-data-dir` に基づく識別名称） |
| `sessionCapabilitiesRequest` | 新しい BiDi-CDP Mapper が起動されたときだけ `session.new` に適用 |
| `WebSocketMode` | WebSocket で制御する場合、接続済みの `CDPCoreViaWebSocket` を指定 |

**戻り値:** 再接続成功可否（`session.status` の ready 判定）。

```vb
' Pipe 版
Dim mode As New WebDriverBiDiMode
If Not mode.reattach(ShSetting01_StartBrowser.CurrentUserName) Then Exit Sub

' WebSocket 版
Dim ws As New CDPCoreViaWebSocket
' ... 接続済み ws を渡す ...
If Not mode.reattach(UserName, , ws) Then Exit Sub
```

::: tip 注意
パイプが生きていない場合は、このメソッドから再開できません。Part1 からやり直してください。
:::

詳細は [再接続](/guides/reattach) / [WebSocket モード](/websocket/capabilities)。

### `quit`

```vb
Public Sub quit()
```

`browser.close` を送り、パイプ／Excel テーブル上のセッション情報を解放します。

```vb
mode.quit
```

## タブ

### `newTab`

```vb
Public Function newTab( _
    Optional newWindow As Boolean, _
    Optional isBackground As Boolean, _
    Optional setMain As Boolean _
) As WebDriverBiDiContext
```

新規タブ（またはウィンドウ）を開き [`WebDriverBiDiContext`](./WebDriverBiDiContext) を返します。BiDi では作成時に URL 直指定はできません。開いたあと `navigate` してください。

| 引数 | 意味 |
| --- | --- |
| `newWindow` | `True` で新しいウィンドウ、`False`（既定）で既存ウィンドウにタブ |
| `isBackground` | `True` でバックグラウンド（非アクティブ）で開く |
| `setMain` | `True` で Excel に context を記録（reattach 用） |

```vb
Dim tab As WebDriverBiDiContext
Set tab = mode.newTab(setMain:=True)
tab.navigate "https://example.com"

Set tab = mode.newTab(newWindow:=True, isBackground:=True)
```

関連: [マルチタブ](/guides/multi-tab)

### `getTab`

```vb
Public Function getTab( _
    Optional Url As String, _
    Optional maxDepth As Long, _
    Optional setMain As Boolean, _
    Optional doRetrySecond As Double _
) As WebDriverBiDiContext
```

既に開いている browsing context を URL 部分一致などで探し、[`WebDriverBiDiContext`](./WebDriverBiDiContext) として返します。見つからない場合は `Nothing` になり得ます。

| 引数 | 意味 |
| --- | --- |
| `Url` | URL 部分一致。省略時は見つかった先頭（未接続相当）の context |
| `maxDepth` | `browsingContext.getTree` の深さ。`0`（既定）ならトップレベルのみ（iframe 除外） |
| `setMain` | `True` で Excel に context を記録 |
| `doRetrySecond` | 指定秒以内に見つかるまでリトライ。`0`（既定）なら 1 回のみ |

```vb
' 直近のタブへ
Dim tab As WebDriverBiDiContext
Set tab = mode.getTab(setMain:=True)

' URL 部分一致
Set tab = mode.getTab(Url:="example.com")

' iframe まで含めて最大 5 秒リトライ
Set tab = mode.getTab(Url:="https://challenges.cloudflare.com/", maxDepth:=2, doRetrySecond:=5)
```

::: tip 注意
WebDriver BiDi ではタブ名（タイトル）での検索はできません。URL で探してください。
:::

### `serializeMainTab`

```vb
Public Property Get serializeMainTab() As String
Public Property Let serializeMainTab(contextId As String)
```

Excel テーブルにメイン browsing context id を記録／読み取ります。`newTab` / `getTab` の `setMain:=True` が内部でこれを使います。

```vb
Debug.Print mode.serializeMainTab
mode.serializeMainTab = tab.context
```

## プロトコル・イベント

### `ExecuteBiDi` / `ExecuteBiDiAsync`

```vb
Public Function ExecuteBiDi(methodName As String, _
    Optional params As Dictionary, _
    Optional StopBiDiError As Boolean = True) As BiDiCDPJson

Public Function ExecuteBiDiAsync(methodName As String, _
    Optional params As Dictionary, _
    Optional StopError As Boolean = True) As Long
```

セッション／ブラウザ向け BiDi コマンドです。結果を待つか（`ExecuteBiDi`）、待たずに後で確認するか（`ExecuteBiDiAsync`）の 2 種類があります。

| 引数 | 意味 |
| --- | --- |
| `methodName` | メソッド名（例: `"browsingContext.navigate"` / `"browser.close"`） |
| `params` | params の `Dictionary`。省略時は空の `{}` |
| `StopBiDiError` / `StopError` | 失敗時に停止するか。既定は `True` |

`ExecuteBiDiAsync` はコマンド実行時の **id（`Long`）のみ**を返し、結果は待ちません。回収は [`TakeResultBiDi`](#takeresultbidi) で行います。

```vb
Dim params As New Dictionary
params.Add "url", "https://example.com"
params.Add "wait", "complete"
' ※ Context 側の ExecuteBiDi は context を自動付与

Dim result As BiDiCDPJson
Set result = mode.ExecuteBiDi("browser.getUserContexts", New Dictionary)
```

詳細は [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)。

### `TakeEvents`

```vb
Public Sub TakeEvents(Optional StopPipeError As Boolean = True, Optional destruction As Boolean)
```

非同期応答／イベントの吸い上げです。[`TakeResultBiDi`](#takeresultbidi) の前に呼び出す必要があります。

| 引数 | 意味 |
| --- | --- |
| `StopPipeError` | Pipe／WebSocket 障害時に停止するか。既定は `True` |
| `destruction` | 破棄処理向けの内部フラグ（通常は省略） |

```vb
mode.TakeEvents
```

### `TakeResultBiDi`

```vb
Property Get TakeResultBiDi(commandID As Long) As String
```

`ExecuteBiDiAsync` が返した id をキーに、蓄積された実行結果（JSON 文字列）を取り出します。取り出し後は Dictionary から削除されます。結果がまだ無い場合は空文字を返します。

```vb
Dim cmdId As Long
cmdId = mode.ExecuteBiDiAsync("browsingContext.navigate", params)

Do
    mode.TakeEvents
    Dim raw As String
    raw = mode.TakeResultBiDi(cmdId)
    If LenB(raw) Then Exit Do
    DoEvents
Loop
```

### `SetLimitBiDi`

```vb
Property Let SetLimitBiDi(Number As Long)
```

BiDi コマンド結果を Dictionary に溜め込む件数の上限です。デフォルトは **65536 件**です。

上限を超えると、パフォーマンス低下を防ぐため蓄積中の結果履歴が **すべて削除**されます。未回収の `ExecuteBiDiAsync` 結果も消える点に注意してください。

```vb
mode.SetLimitBiDi = 1000
```

::: tip
コマンド ID がオーバーフロー対策でリセットされるとき（およそ 20 億到達時）も、結果履歴はすべてクリアされます。
:::

### `LastBiDiJsonError`

```vb
Property Get LastBiDiJsonError() As Dictionary
```

`ExecuteBiDi` 経由で記録された、**最後の BiDi コマンドエラー**です。`Err.LastDllError` と同様、成功しても消えません。

`StopBiDiError:=False` のとき、戻り値が `Nothing` なら本プロパティで詳細を確認します。

```vb
Dim result As BiDiCDPJson
Set result = mode.ExecuteBiDi("webExtension.install", params, False)

If result Is Nothing Then
    Debug.Print mode.LastBiDiJsonError("message")
    ' 必要なら mode.LastBiDiJsonError.RemoveAll でクリア可
End If
```

### `BiDiEvents`

```vb
Property Get BiDiEvents() As Dictionary
Property Set BiDiEvents(ObjDic As Dictionary)
```

標準モジュール上で BiDi の非同期イベントを受け取るための蓄積口です（CDP の `BrowserEvents` 相当）。

| 操作 | 意味 |
| --- | --- |
| `Set … = New Dictionary` | 記録開始 |
| `Set … = Nothing` | 記録停止 |
| 退避した Dictionary を再代入 | セーブ／再開 |

```vb
Set mode.BiDiEvents = New Dictionary
' ... sessionSubscribe 後に操作 ...
mode.TakeEvents
' mode.BiDiEvents を参照
Set mode.BiDiEvents = Nothing
```

### `sessionSubscribe`

```vb
Property Set sessionSubscribe(Optional subscribe As Boolean = True, events As Collection)
```

`session.subscribe` / `session.unsubscribe` を実行します。どのイベントを購読中かの管理は呼び出し側で行います。

| 引数 | 意味 |
| --- | --- |
| `subscribe` | `True`（既定）で購読、`False` で購読解除 |
| `events` | イベント名の `Collection`（例: `"network.beforeRequestSent"`） |

```vb
Dim events As New Collection
events.Add "network.beforeRequestSent"
events.Add "network.responseCompleted"
events.Add "log.entryAdded"
Set mode.sessionSubscribe = events

' 解除
Set mode.sessionSubscribe(False) = events
```

手順・セーブ／再開は [イベント購読](/guides/events) を参照してください。

## タイムアウト

### `TimeOutSecond`

```vb
Property Get TimeOutSecond() As Double
Property Let TimeOutSecond(TimeSec As Double)
```

BiDi コマンド結果待ちの上限です。デフォルトは **30 秒**です。

```vb
mode.TimeOutSecond = 60
```

タブ側（[`WebDriverBiDiContext`](./WebDriverBiDiContext)）からは `InheritanceWebDriverBiDiMode.TimeOutSecond` で同じ値を触れます。

詳細は [タイムアウト設定方法について](/guides/timeout)。

## ユーティリティ

### `sleep`

```vb
Public Sub sleep(Optional seconds As Double = 0.5)
```

指定秒待ちます。既定は `0.5` 秒です。

```vb
mode.sleep 1
```

### `TimerCounter`

```vb
Public Function TimerCounter() As Double
```

単調増加の経過ミリ秒です。VBA の `Timer` 関数の代わりに、自前ループのタイムアウト判定へ使えます。

```vb
Dim startMs As Double
startMs = mode.TimerCounter
Do
    mode.TakeEvents
    If mode.TimerCounter - startMs > 5000 Then Exit Do
    DoEvents
Loop
```

詳細は [タイムアウト設定方法について](/guides/timeout)。

### `printMsg`

```vb
Public Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, _
    Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
```

デバッグ／ログ出力です。通常はフレームワーク内部から呼ばれます。

## 高度な／内部寄りの API

日常利用では意識不要です。拡張やフレームワーク連携向けです。

### `RunCollect_InfoList`

```vb
Property Set RunCollect_InfoList(ArgCollect As Collection)
```

`browsingContext.contextCreated` で得た InfoList を、渡した `Collection` に蓄積します。`Nothing` で収集停止（用が済んだら必ず解放）。

```vb
Dim info As New Collection
Set mode.RunCollect_InfoList = info
' ... タブ作成などの操作 ...
Set mode.RunCollect_InfoList = Nothing
```

### `EnableDiscoverContexts`

```vb
Property Let EnableDiscoverContexts(Flag As Boolean)
```

`browsingContext.contextCreated` / `contextDestroyed` の購読と、確保中 `WebDriverBiDiContext` のカウントに使います。Context の生成／破棄時に呼ばれる想定です。

### `InheritanceBiDiCore`

```vb
Property Get InheritanceBiDiCore() As WebDriverBiDiCore
```

内部の `WebDriverBiDiCore` への参照です。通常は [`WebDriverBiDiContext`](./WebDriverBiDiContext) 経由で十分です。

## 関連

- [`WebDriverBiDiContext`](./WebDriverBiDiContext)
- [CDP と BiDi](/concepts/cdp-vs-bidi)
- [マルチタブ](/guides/multi-tab)
- [イベント購読](/guides/events)
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)
- [タイムアウト設定方法について](/guides/timeout)
- [再接続](/guides/reattach)
