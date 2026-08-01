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

Public Function ExecuteBiDiAsync(...) As Long
```

BiDi コマンドです。結果を待つか（`ExecuteBiDi`）、待たずに後で確認するか（`ExecuteBiDiAsync`）の 2 種類があります。

`ExecuteBiDiAsync` はコマンド実行時の **id（`Long`）のみ**を返し、結果は待ちません。結果の回収は [`TakeResultBiDi`](#takeresultbidi) で行います。

[低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)

### `TakeEvents`

非同期応答／イベントの吸い上げ。`TakeResultBiDi` の前に呼び出す必要があります。

### `TakeResultBiDi`

```vb
Property Get TakeResultBiDi(commandID As Long) As String
```

`ExecuteBiDiAsync` が返した id をキーに、蓄積された実行結果（JSON 文字列）を取り出します。取り出し後は Dictionary から削除されます。結果がまだ無い場合は空文字を返します。

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

### `sessionSubscribe` / `BiDiEvents`

イベント購読の中核。[イベント購読](/guides/events)

### `LastBiDiJsonError`

`StopBiDiError:=False` 時のエラー情報。

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

| メンバー | 説明 |
| --- | --- |
| `sleep` | 秒待ち |
| `TimerCounter` | 経過ミリ秒。`Timer` 関数の代わりに自前ループのタイムアウト判定へ。[タイムアウト設定方法について](/guides/timeout) |
| `printMsg` | デバッグ出力 |

## 関連

- [`WebDriverBiDiContext`](./WebDriverBiDiContext)
- [CDP と BiDi](/concepts/cdp-vs-bidi)
- [マルチタブ](/guides/multi-tab)
- [タイムアウト設定方法について](/guides/timeout)
