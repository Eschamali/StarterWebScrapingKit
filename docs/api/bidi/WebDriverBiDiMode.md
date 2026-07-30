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

タブからすぐ始めたい場合は `StartBiDiModeContext` → [`WebDriverBiDiContext`](./WebDriverBiDiContext)。

## 起動・再接続・終了

### `StartBiDiMode`

```vb
Public Sub StartBiDiMode(Optional Name As String = "chrome", ...)
```

低レベル起動。通常は設定シート経由。

### `reattach`

```vb
Public Function reattach(userProfile As String, _
    Optional sessionCapabilitiesRequest As Dictionary, _
    Optional WebSocketMode As CDPCoreViaWebSocket) As Boolean
```

[再接続](/guides/reattach)

### `quit`

セッションとブラウザを終了。

## タブ

### `newTab`

```vb
Public Function newTab(Optional newWindow As Boolean, Optional isBackground As Boolean, _
    Optional setMain As Boolean) As WebDriverBiDiContext
```

### `getTab`

```vb
Public Function getTab(Optional Url As String, Optional maxDepth As Long, _
    Optional setMain As Boolean, Optional doRetrySecond As Double) As WebDriverBiDiContext
```

URL 部分一致などで検索。見つからない場合は `Nothing` になり得ます。

### `serializeMainTab`

メイン browsing context id の記録／取得。

## プロトコル・イベント

### `ExecuteBiDi` / `ExecuteBiDiAsync`

```vb
Public Function ExecuteBiDi(methodName As String, _
    Optional params As Dictionary, _
    Optional StopBiDiError As Boolean = True) As BiDiCDPJson
```

[低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)

### `sessionSubscribe` / `BiDiEvents` / `TakeEvents`

イベント購読の中核。[イベント購読](/guides/events)

### `LastBiDiJsonError`

`StopBiDiError:=False` 時のエラー情報。

## ユーティリティ

`sleep` / `TimerCounter`（`Timer` 代替・自前ループ用） / `printMsg`

詳細は [タイムアウト設定方法について](/guides/timeout)。

## 関連

- [`WebDriverBiDiContext`](./WebDriverBiDiContext)
- [CDP と BiDi](/concepts/cdp-vs-bidi)
- [マルチタブ](/guides/multi-tab)
