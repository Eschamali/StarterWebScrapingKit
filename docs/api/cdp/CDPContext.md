# CDPContext

1 つのタブ（ページ）を表します。Playwright の **Page** に相当します。

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://example.com"
t.InheritanceCDPBrowser.quit
```

親ブラウザは `InheritanceCDPBrowser`（[`CDPBrowser`](./CDPBrowser)）です。日常利用では設定シート経由の `StartCDPModeContext` を推奨します。

## 起動・再接続・終了

### `StartAndConnectTab`

ブラウザ起動と同時に、初回タブ接続まで自動で行います。

```vb
Public Sub StartAndConnectTab( _
    Optional Name As String = "chrome", _
    Optional appUrl As String, _
    Optional userProfile As String, _
    Optional addArgs As String, _
    Optional KioskMode As edgeKioskType _
)
```

| 引数 | 意味 |
| --- | --- |
| `Name` | ブラウザ名（現時点では Chrome / Edge） |
| `appUrl` | `--app` に付ける URL |
| `userProfile` | `--user-data-dir` 用のユーザーディレクトリ名 |
| `addArgs` | 追加の起動引数 |
| `KioskMode` | Edge キオスクモード |

日常利用では設定シート経由で十分です。低レベルに起動したいときだけ直接呼んでください。

### `reattach`

Excel テーブルにある既存のパイプハンドル／ブラウザセッション情報を利用して、再接続を試みます。

```vb
Public Function reattach( _
    userProfile As String, _
    Optional reuseSession As Boolean, _
    Optional WebSocketMode As CDPCoreViaWebSocket _
) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `userProfile` | 再アタッチしたいユーザー名（`user-data-dir` に基づく） |
| `reuseSession` | `True` で Excel に記録中の SessionID を流用。`False` なら `targetID` を基に SessionID を更新し、古い SessionID を破棄して上書き |
| `WebSocketMode` | WebSocket で CDP 制御する場合、接続処理済みの `CDPCoreViaWebSocket` を指定 |

**戻り値:** 既存タブへの接続成功可否。

```vb
' Pipe 版
Dim t As New CDPContext
If Not t.reattach(ShSetting01_StartBrowser.CurrentUserName) Then Exit Sub

' SessionID を引き継ぐ（KeepSession 済みの場合）
If Not t.reattach(UserName, True) Then Exit Sub

' WebSocket 版（Page 接続）
Dim ws As New CDPCoreViaWebSocket
ws.AutoConnectPageCDP UserName
If Not t.reattach(UserName, , ws) Then Exit Sub
```

::: tip 注意
- この処理は Excel テーブルに記録されている `targetID` への再接続までです。`targetID` 自体が既に閉じている場合は `False` になりますが、パイプが生きていれば親の [`newTab` / `getTab`](./CDPBrowser)（`setMain:=True`）で手動復旧できます
- 非同期イベント購読中（例: `Network.webSocketFrameReceived`）は `reuseSession:=True` を推奨します
:::

詳細は [再接続ガイド](/guides/reattach)（SessionID 引き継ぎ含む）/ [WebSocket モード](/websocket/capabilities) を参照。

### `closeTab`

```vb
Public Sub closeTab()
```

このタブを閉じます。ブラウザ全体を終了する [`CDPBrowser.quit`](./CDPBrowser#quit) とは別です。

::: tip 注意
メインタブ（Excel テーブルに記録されたセッションタブ）は閉じられません。
:::

## ナビ・待機・表示

### `navigate`

```vb
Public Sub navigate(strURL As String, Optional till As ReadyState = isComplete)
```

URL を開き、指定の [`ReadyState`](#readystate) まで待ちます。内部では `Page.navigate` のあと [`wait`](#wait) を呼びます。

| 引数 | 意味 |
| --- | --- |
| `strURL` | 遷移先 URL |
| `till` | 待機する `document.readyState`。既定は `isComplete`（読み込み完了） |

```vb
t.navigate "https://example.com"                          ' 完了まで待つ（既定）
t.navigate "https://example.com/heavy", isInteractive     ' interactive で先に進む
```

::: tip 注意
すでに同じ URL にいる場合は遷移せず、警告ログを出して終了します。
:::

### `wait`

```vb
Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
```

現在ページが指定の [`ReadyState`](#readystate) になるまで待ちます。`navigate` 後だけでなく、クリック後の再読み込み待ちなどにも使えます。

| 引数 | 意味 |
| --- | --- |
| `till` | 待機する `document.readyState`。既定は `isComplete` |
| `dbgState` | `True` で待機中の ReadyState を Immediate ウィンドウへ出し、`jsEval` の例外も止めずに継続 |

```vb
t.wait                              ' 完了待ち（既定）
t.wait isInteractive                ' interactive で十分なら短縮できる
t.wait isComplete, dbgState:=True   ' 状態遷移を見ながら待つ
```

::: tip
`till:=isInteractive` のとき、すでに `complete` まで進んでいればそのまま成功扱いで抜けます（interactive を取りこぼしても止まらない）。
:::

### `ReadyState`

`document.readyState` に対応する列挙です。`navigate` / `wait` / 一部の要素操作で使います。

| 値 | ブラウザ側 | 意味 |
| --- | --- | --- |
| `isLoading` | `"loading"` | 読み込み中 |
| `isInteractive` | `"interactive"` | DOM は操作可能だが、画像などの読み込みは未完了のことがある |
| `isComplete` | `"complete"` | ドキュメント読み込み完了（既定）。要素が完了後にしか出ないページ向け |

詳細は [ページ遷移](/guides/navigation)。

### `show` / `hide` / `activate` / `bringToForeground`

ウィンドウ表示制御。`show` は `xywh:="0 20 1000 700"` のような配置も可。

## JavaScript・通知

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, Optional objectId As String, _
    Optional objectArguments As Collection, Optional IFEXCEPTION As Variant, ...) As Variant
```

詳細は [JavaScript 実行](/guides/javascript)。

### `jsAddLib` / `jsAddScript`

外部／ローカルスクリプトの注入。

### `notify`

```vb
Public Sub notify(msg As String, Optional ViewSecond As Long = 10)
```

ページ上に一時メッセージ。

### `handleDialog`

```vb
Public Sub handleDialog(Accept As Boolean, Optional promptText As String)
```

alert / confirm / prompt への応答。

## 要素

| メソッド | 説明 |
| --- | --- |
| `getElementByID` | id |
| `getElementByQuery` / `getElementsByQuery` | CSS |
| `getElementByXPath` / `getElementsByXPath` | XPath |

戻り値は [`CDPElement`](./CDPElement)。詳細は [要素の取得](/guides/selectors)。

## スクリーンショット

### `snapPage`

```vb
Public Sub snapPage(FolderPath As String, FileName As String, Optional getFullPage As Boolean = False)
```

[スクリーンショット](/guides/screenshots)

## プロトコル・イベント

### `ExecuteCDP` / `ExecuteCDPAsync`

ページ／セッション向け CDP。結果を待つか（`ExecuteCDP`）、待たずに後で確認するか（`ExecuteCDPAsync`）の 2 種類があります。

`ExecuteCDPAsync` はコマンド実行時の **id（`Long`）のみ**を返し、結果は待ちません。結果の回収は [`TakeResultCDP`](#takeresultcdp) で行います。

詳細は [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol) で説明します。

### `TakeResultCDP`

```vb
Property Get TakeResultCDP(commandID As Long) As String
```

`ExecuteCDPAsync` が返した id をキーに、蓄積された実行結果（JSON 文字列）を取り出します。取り出し前に `TakeEvents` が必要です。取り出し後は Dictionary から削除され、結果がまだ無い場合は空文字を返します。

### `SetLimitCDPResult`

```vb
Property Let SetLimitCDPResult(Number As Long)
```

CDP コマンド結果を Dictionary に溜め込む件数の上限です。デフォルトは **65536 件**です。

上限を超えると、パフォーマンス低下を防ぐため蓄積中の結果履歴が **すべて削除**されます。未回収の `ExecuteCDPAsync` 結果も消える点に注意してください。

```vb
t.SetLimitCDPResult = 1000
```

::: tip
コマンド ID がオーバーフロー対策でリセットされるとき（およそ 20 億到達時）も、結果履歴はすべてクリアされます。
:::

### `BrowserEvents` / `SetFilterEvents`

イベント蓄積用 Dictionary とフィルタ。[イベント購読](/guides/events)

### `pageEnable` / `runtimeEnable`

ドメイン有効化のショートカット。

### `openDevTools`

このタブの DevToolsを開きます。

> [!WARNING]
> WebView2/Electron製の場合は、うまくいかない場合があります


## タイムアウト

### `TimeOutSecond`

```vb
Property Get TimeOutSecond() As Double
Property Let TimeOutSecond(TimeSec As Double)
```

タブ単位の CDP コマンド結果待ちや、起動直後の遷移完了判定などの待機上限です。デフォルトは **30 秒**です。

```vb
t.TimeOutSecond = 60
```

自前ループ用の経過ミリ秒は親の [`CDPBrowser.TimerCounter`](./CDPBrowser) を使います。詳細は [タイムアウト設定方法について](/guides/timeout)。

## デバッグ

`printParams` / `getSessionInfo` / `printMsg`

## 関連

- [`CDPBrowser`](./CDPBrowser)
- [`CDPElement`](./CDPElement)
- [はじめに](/getting-started)
- [ページ遷移](/guides/navigation)
- [タイムアウト設定方法について](/guides/timeout)
