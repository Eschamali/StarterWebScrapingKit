# CDPContext

1 つのタブ（ページ）を表します。Playwright の **Page** に相当します。

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://example.com"
t.InheritanceCDPBrowser.quit
```

親ブラウザは `InheritanceCDPBrowser`（`CDPBrowser`）です。

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

### `closeTab`

タブを閉じます（ブラウザ全体の `quit` とは別）。

## JavaScript・通知

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, Optional objectId As String, _
    Optional objectArguments As Collection, Optional IFEXCEPTION As Variant, ...) As Variant
```

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

ページ／セッション向け CDP。[生プロトコル拡張](/guides/extend-raw-protocol)

### `BrowserEvents` / `SetFilterEvents`

イベント蓄積用 Dictionary とフィルタ。[イベント購読](/guides/events)

### `pageEnable` / `runtimeEnable`

ドメイン有効化のショートカット。

### `openDevTools`

このタブの DevTools。

## 再接続

### `reattach`

```vb
Public Function reattach(userProfile As String, Optional reuseSession As Boolean, _
    Optional WebSocketMode As CDPCoreViaWebSocket) As Boolean
```

[再接続](/guides/reattach)

### `StartAndConnectTab`

低レベル起動（通常は設定シート経由で十分）。

## デバッグ

`printParams` / `getSessionInfo` / `printMsg`

## 関連

- [`CDPBrowser`](./CDPBrowser)
- [`CDPElement`](./CDPElement)
- [はじめに](/getting-started)
