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

URL を開き、指定 ReadyState まで待ちます。

### `wait`

```vb
Public Sub wait(Optional till As ReadyState = isComplete, Optional dbgState As Boolean = False)
```

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

ページ／セッション向け CDP。[低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)

### `BrowserEvents` / `SetFilterEvents`

イベント蓄積用 Dictionary とフィルタ。[イベント購読](/guides/events)

### `pageEnable` / `runtimeEnable`

ドメイン有効化のショートカット。

### `openDevTools`

このタブの DevToolsを開きます。

> [!WARNING]
> WebView2/Electron製の場合は、うまくいかない場合があります


## デバッグ

`printParams` / `getSessionInfo` / `printMsg` 

## 関連

- [`CDPBrowser`](./CDPBrowser)
- [`CDPElement`](./CDPElement)
- [はじめに](/getting-started)
- [タイムアウト設定方法について](/guides/timeout)
