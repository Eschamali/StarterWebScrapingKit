---
description: タブ単位のナビ・ウィンドウ制御・jsEval・iframe・イベント・スクリーンショットなど、CDP のページ操作を網羅します。
---

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

## ナビ・待機

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

## ウィンドウ制御

### `show`

```vb
Public Sub show(Optional state As WinState = asNormal, Optional xywh As String = "0 0 0 0")
```

ブラウザウィンドウを表示し、必要ならサイズ／位置を変更します。[`WinState`](#winstate)（[ShowWindow](https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow) 準拠）と、CDP の `Browser.setWindowBounds` を組み合わせています。reattach 後にウィンドウを前面へ出す用途にも使えます。

| 引数 | 意味 |
| --- | --- |
| `state` | 表示状態。既定は `asNormal`（通常表示＋アクティブ化） |
| `xywh` | `"left top width height"` 形式の文字列。省略時（`"0 0 0 0"`）はリサイズしない |

```vb
t.show asMaximized
t.show asNormal, "100 50 1200 800"   ' 左上とサイズを指定
t.show , "100 200"                  ' 位置だけ変更（幅・高さは据え置き）
```

::: tip 注意
- 最大化中のリサイズは意図どおりに効かないことがあります。先に `asNormal` などで戻してから `xywh` を指定してください
- `xywh` の各値が `0` の項目はスキップされます（例: `"100 200"` は left / top のみ）
:::

### `hide`

```vb
Public Sub hide()
```

ウィンドウを非表示にします（`ShowWindow` の `asHidden` 相当）。`show` で再度表示できます。

```vb
t.hide
' ... 裏で処理 ...
t.show
```

### `activate`

```vb
Public Sub activate()
```

**ブラウザ内のタブ**にフォーカスを移します（CDP の `Target.activateTarget`）。ウィンドウ全体を前面に出すわけではない点で、`show` / `bringToForeground` とは役割が違います。

```vb
Dim tab2 As CDPContext
Set tab2 = t.InheritanceCDPBrowser.getTab(Url:="example.com")
tab2.activate   ' そのタブを前面タブにする
```

### `bringToForeground`

```vb
Public Function bringToForeground()
```

ブラウザ**ウィンドウ**を最前面にします（`ShowWindow` → `BringWindowToTop` → `SetForegroundWindow`）。OS のフォーカス制限により、常に最前面になるとは限りません。

```vb
t.bringToForeground
```

### `BrowserWindowHandle` / `BrowserWindowID`

```vb
Property Get BrowserWindowHandle(Optional alwaysRequest As Boolean) As LongPtr
Property Get BrowserWindowID(Optional alwaysRequest As Boolean) As Long
```

ウィンドウ操作の土台になる ID です。`show` / `hide` / `bringToForeground` は内部でこれらを使います。自前で WinAPI や CDP のウィンドウ調整をするときにも参照できます。

| プロパティ | 意味 | 主な用途 |
| --- | --- | --- |
| `BrowserWindowHandle` | OS のウィンドウハンドル（`HWND`） | `ShowWindow` など WinAPI |
| `BrowserWindowID` | CDP の `windowId`（`Browser.getWindowForTarget`） | `Browser.setWindowBounds` など CDP |

どちらも `alwaysRequest:=False`（既定）ならキャッシュがあれば流用し、無ければ調査します。`True` なら毎回取り直します。

```vb
Dim hwnd As LongPtr
hwnd = t.BrowserWindowHandle

Dim wid As Long
wid = t.BrowserWindowID
```

### `WinState`

`show` の第 1 引数に渡す列挙です。[ShowWindow の nCmdShow](https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow) に準拠しています。よく使うのは次のとおりです。

| 値 | 意味 |
| --- | --- |
| `asNormal` | 通常表示し、アクティブ化（既定）。最小化／最大化なら元のサイズへ |
| `asMinimized` | 最小化してアクティブ化 |
| `asMaximized` | 最大化してアクティブ化 |
| `doShowNoActivate` | 表示するがアクティブ化しない |
| `doShowMinNoActivate` | 最小化表示するがアクティブ化しない |
| `asHidden` | 非表示（通常は [`hide`](#hide) を使う） |

その他（`doShow` / `doRestore` / `doForceMin` など）も Enum に定義されています。起動時の初期表示モード設定でも同じ列挙を使います（[はじめに](/getting-started)）。

## JavaScript・通知

### `jsEval`

```vb
Public Function jsEval(JavaScriptStr As String, Optional objectId As String, _
    Optional objectArguments As Variant, Optional IFEXCEPTION As Variant, ...) As Variant
```

ページ上で JavaScript を評価します。例外時は `IsError(result)` で判定し、詳細は [`LastJavaScriptException`](#lastjavascriptexception) を参照してください。代替値で済ませたい場合は `IFEXCEPTION` を使います。

`objectArguments` は `Collection` / `Array(...)` / 固定長の `Dictionary` 型 1 次元配列（`Dim args(0) As Dictionary`）のいずれでも渡せます。**固定長配列の方がパフォーマンスが良い**ため、ループ内などで多用する場合は固定長配列を推奨します。

詳細は [JavaScript 実行](/guides/javascript)。

::: tip
`objectId` / `contextId` のいずれも省略したとき（このタブの既定コンテキストで実行するケース）、内部で保持する `executionContextId` がページ遷移等で無効化されていた場合は、有効な ID が届くまで自動で待機してから実行します。
:::

### `LastJavaScriptException`

```vb
Property Get LastJavaScriptException() As BiDiCDPJson
```

同期の `jsEval` で起きた、**最後の JavaScript 例外**（CDP の `exceptionDetails`）です。`Err.LastDllError` と同様、成功しても消えません。

```vb
Dim result As Variant
result = t.jsEval("notDefined.x")

If IsError(result) Then
    Debug.Print t.LastJavaScriptException.Stringify
    ' 必要なら .StringKey / .NodeKey で個別フィールドも参照可
End If

' 詳細を見なくてよいなら代替値で済ませる
result = t.jsEval("notDefined.x", IFEXCEPTION:="fallback")
```

::: tip 注意
- 対象は **例外**のみです。`try { ... } catch (e) { return e }` のように JS 側で捕まえて返した値は対象外です
- 非同期（`RunAsyncCDP:=True`）で後から取り出した結果とは連動しません
- JS は成功したが戻り値がエラー値の場合は `IsError` で対処してください（`IFEXCEPTION` / 本プロパティは「例外」時のみ）
:::

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

## iframe

ページに埋め込まれた iframe へ入るための一連の API です。流れは **一覧 → `executionContextId` 取得 → `CDPElement` 化** です。

同一オリジンの `<iframe>` 要素から直接入りたい場合は、要素側の [`CDPElement.getIFrame`](./CDPElement#getiframe) もあります。

### `printChildFrames`

```vb
Public Sub printChildFrames()
```

現時点で開いている **直下の子 iframe** 情報をログに出します（`Page.getFrameTree`）。名前・URL・frameID などを確認するときに使います。

```vb
t.printChildFrames
' Immediate ウィンドウに title / url / frameID などが並ぶ
```

::: tip 注意
検索範囲は **親の直下の子 iframe まで**です。子 iframe の中の iframe… といった多重ネスト表示には対応していません。
:::

### `getIFrameContextID`

```vb
Public Function getIFrameContextID( _
    Optional iframeName As String, _
    Optional Url As String, _
    Optional doRetrySecond As Double _
) As Long
```

埋め込み iframe の `executionContextId` を返します。[`jsEval`](#jseval) の `contextId` に渡せます。見つからない場合は `0` です。

| 引数 | 意味 |
| --- | --- |
| `iframeName` | iframe 名（第一優先。`<iframe name="…">`。完全一致 → 部分一致） |
| `Url` | URL（第二優先。先頭一致 → 部分一致） |
| `doRetrySecond` | 指定秒以内に見つかるまでリトライ。`0`（既定）なら 1 回のみ |

```vb
' 名前で探す
Dim ctxId As Long
ctxId = t.getIFrameContextID(iframeName:="app-frame")

' URL 部分一致 + 最大 5 秒リトライ
ctxId = t.getIFrameContextID(Url:="https://example.com/embed", doRetrySecond:=5)

' 両方省略 → 最初の iframe
ctxId = t.getIFrameContextID()

' jsEval で直接使う
Debug.Print t.jsEval("document.title", contextId:=ctxId)
```

::: tip 注意
- `iframeName` と `Url` を両方省略すると、最初の iframe の `executionContextId` を返します
- `CDPElement` として扱いたい場合は [`getIFrame`](#getiframe) に渡してください
- 検索範囲は **親の直下の子 iframe まで**です。多重ネストは非対応です。その場合は自力で `ExecuteCDP("Page.getFrameTree").NodeKey("frameTree")` から目的の frame を取り出してください
:::

### `getIFrame`

```vb
Public Function getIFrame(ExecutionContextId As Long) As CDPElement
```

[`getIFrameContextID`](#getiframecontextid) で得た `executionContextId` を、[`CDPElement`](./CDPElement) として扱えるようにします（その context 上の `document`）。

| 引数 | 意味 |
| --- | --- |
| `ExecutionContextId` | `getIFrameContextID` が返した `executionContextId` |

```vb
Dim ctxId As Long
ctxId = t.getIFrameContextID(iframeName:="app-frame")

Dim frame As CDPElement
Set frame = t.getIFrame(ctxId)
frame.getElementByQuery("button").click
```

## スクリーンショット

### `snapPage`

```vb
Public Sub snapPage(FolderPath As String, FileName As String, Optional getFullPage As Boolean = False)
```

現在のタブを PNG として保存します。内部では CDP の `Page.captureScreenshot` を使います（外部 JS ライブラリは不要です）。

| 引数 | 意味 |
| --- | --- |
| `FolderPath` | 保存先フォルダ（例: `Environ("UserProfile") & "\Downloads"`） |
| `FileName` | ファイル名（拡張子込み。例: `"shot.png"`） |
| `getFullPage` | `False`（既定）で現在のビューポートのみ。`True` でページ全体（縦スクロール範囲を含む） |

```vb
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://example.com"

' ビューポートのみ
t.snapPage Environ("UserProfile") & "\Downloads", "viewport.png"

' フルページ
t.snapPage Environ("UserProfile") & "\Downloads", "full.png", True

t.InheritanceCDPBrowser.quit
```

::: tip 注意
- レイアウト情報（`Page.getLayoutMetrics`）が取れない画面では警告を出して終了します
:::

デモ: `Demo_CDP.getSnapShot`

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

### `LastCDPJsonError`

```vb
Property Get LastCDPJsonError() As Dictionary
```

`ExecuteCDP` 経由で記録された、**最後の CDP コマンドエラー**です。`Err.LastDllError` と同様、成功しても消えません。

`StopCDPError:=False` で例外停止を抑えたとき、戻り値が `Nothing` なら本プロパティで詳細を確認します。キーは主に `"code"` / `"message"` です。

```vb
Dim result As BiDiCDPJson
Set result = t.ExecuteCDP("Page.navigate", params, False)

If result Is Nothing Then
    Debug.Print t.LastCDPJsonError("message")
    ' 必要なら t.LastCDPJsonError.RemoveAll でクリア可
End If
```

::: tip 注意
`ExecuteError` を通らない経路では更新されません。目印は「`ExecuteCDP` の戻り値が `Nothing`」です。
:::

### `BrowserEvents` / `SetFilterEvents`

```vb
Property Get BrowserEvents() As Dictionary
Property Set BrowserEvents(ObjDic As Dictionary)

Property Let SetFilterEvents(Optional DelMode As Boolean, EventName As String)
```

標準モジュール上で CDP の非同期イベントを受け取るための蓄積口です。

| メンバー | 役割 |
| --- | --- |
| `BrowserEvents` | `New Dictionary` を渡すと記録開始、`Nothing` で停止。退避した Dictionary を戻せば再開も可 |
| `SetFilterEvents` | 蓄積するイベント名（例: `"Network.requestWillBeSent"`）を絞る。未設定時は広くキャプチャ |

記録開始後は、対象ドメインを `ExecuteCDP` で enable してから操作します。

```vb
t.SetFilterEvents = "Network.requestWillBeSent"
Set t.BrowserEvents = New Dictionary
t.ExecuteCDP "Network.enable"
' ... 操作後、t.BrowserEvents を参照 ...
Set t.BrowserEvents = Nothing
```

手順・セーブ／再開・`WithEvents` との使い分けは [イベント購読](/guides/events) を参照してください。

### `pageEnable` / `runtimeEnable`

ドメイン有効化のショートカット。

### `openDevTools`

このタブの DevToolsを開きます。

> [!WARNING]
> WebView2/Electron製の場合は、うまくいかない場合があります


## タイムアウト

### `TimeOutSecond`

```vb
Property Let TimeOutSecond(TimeSec As Double)
```

タブ単位の CDP コマンド結果待ちや、起動直後の遷移完了判定などの待機上限です。デフォルトは **30 秒**です。**LET 専用**（書き込みのみ）で、設定中の値は読み返せません。

```vb
t.TimeOutSecond = 60
```

自前ループ用の経過ミリ秒は親の [`CDPBrowser.TimerCounter`](./CDPBrowser) を使います。詳細は [タイムアウト設定方法について](/guides/timeout)。

## タブ情報

### `Url`

```vb
Property Get Url() As String
Property Let Url(newURL As String)
```

現在タブの URL です。代入すると内部で [`navigate`](#navigate) を呼びます。

```vb
Debug.Print t.Url
t.Url = "https://example.com"   ' t.navigate "https://example.com" と同じ
```

### `Title`

```vb
Property Get Title() As String
Property Let Title(newName As String)
```

タブのタイトル（`document.title` 相当）です。`jsEval("document.title")` で取得します。代入するとタイトルを書き換えます。

```vb
Debug.Print t.Title
t.Title = "作業用タブ"
```

### `html`

```vb
Property Get html() As String
```

ページ全体の HTML 文字列です。内部では `jsEval("document.documentElement.innerHTML")` を実行します（Get 専用）。

```vb
Debug.Print Left$(t.html, 200)   ' 先頭だけ確認する例
```

巨大なページでは文字列が長くなる点に注意してください。部分だけ欲しい場合は [`jsEval`](#jseval) でセレクタ付きに取る方が軽くて済みます。

### `CurrentSessionID` / `CurrentTargetID` / `CurrentBrowserContextId`

```vb
Property Get CurrentSessionID() As String
Property Get CurrentTargetID() As String
Property Get CurrentBrowserContextId() As String
```

このタブ接続を識別する ID です。

| プロパティ | 意味 |
| --- | --- |
| `CurrentSessionID` | CDP セッション ID。コマンド送信先のセッションを表す |
| `CurrentTargetID` | ターゲット ID。「どのタブか」を特定するための専用 ID |
| `CurrentBrowserContextId` | 「どの Profile（ブラウザコンテキスト）に属しているか」を特定するための専用 ID |

```vb
Debug.Print t.CurrentSessionID
Debug.Print t.CurrentTargetID
Debug.Print t.CurrentBrowserContextId
```

[再接続 (reattach)](/guides/reattach) やメインタブ記録で Excel テーブルに残る値と対応します。日常のページ操作では意識不要です。

`CurrentBrowserContextId` は、`Target.createBrowserContext` で作成した独立プロファイル（シークレット相当）を複数並行運用するときに使います。

## タブ生存について

CDP のターゲット／セッションイベントを受けたあと、このタブがまだ使えるかを確認するフラグです。いずれも **Get 専用**で、該当イベントを受け取ると `True` になります。

```vb
Property Get isTargetDisconnected() As Boolean
Property Get isTargetCrashed() As Boolean
Property Get isTargetClosed() As Boolean
```

| プロパティ | 対応イベント | 意味 |
| --- | --- | --- |
| `isTargetDisconnected` | `Target.detachedFromTarget` | このタブの **セッション**が切れた。`targetId` 自体は残っていることがある。セッションの更新（再接続など）が必要 |
| `isTargetCrashed` | `Target.targetCrashed` | タブがクラッシュした。`Page.reload` で復帰できる場合がある |
| `isTargetClosed` | `Target.targetDestroyed` | タブが閉じられた。この `CDPContext` はもう使えないので破棄する |

```vb
If t.isTargetClosed Then
    ' オブジェクトを捨てて、必要なら getTab / newTab で取り直す
ElseIf t.isTargetDisconnected Then
    ' reattach などでセッションを更新
ElseIf t.isTargetCrashed Then
    t.ExecuteCDP "Page.reload"   ' 復帰を試す例
End If
```

::: tip 注意
`isTargetClosed`（`Target.targetDestroyed`）の検知には、ブラウザ側で `Target.setDiscoverTargets` の購読が必要です。通常の起動フローでは内部で有効化されます。
:::

関連: [再接続 (reattach)](/guides/reattach)

## 親ブラウザ

### `InheritanceCDPBrowser`

```vb
Property Get InheritanceCDPBrowser() As CDPBrowser
```

`CDPContext` 上で、ブラウザ単位の制御もしたいときに使います。親の [`CDPBrowser`](./CDPBrowser) への参照です。

タブ操作（`navigate` など）は Context 側、プロセス／タブ一覧／ブラウザ向け CDP（`quit` / `newTab` / `ExecuteCDP` など）はこちら経由、という使い分けになります。

```vb
' 別タブを開く
t.InheritanceCDPBrowser.newTab "https://example.com"

' ブラウザ向け CDP（拡張機能の読み込みなど）
t.InheritanceCDPBrowser.ExecuteCDP "Extensions.loadUnpacked", params

' ブラウザ終了
t.InheritanceCDPBrowser.quit
```

## デバッグ

`printParams` / `getSessionInfo` / `printMsg`

## 関連

- [`CDPBrowser`](./CDPBrowser)
- [`CDPElement`](./CDPElement)
- [はじめに](/getting-started)
- [ページ遷移](/guides/navigation)
- [要素の取得](/guides/selectors)
- [タイムアウト設定方法について](/guides/timeout)
