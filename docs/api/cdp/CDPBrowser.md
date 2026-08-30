---
description: ブラウザ起動・再接続・タブ操作・ExecuteCDP など、プロセス単位の CDP 制御インターフェースを解説します。
---

# CDPBrowser

ブラウザプロセス単位のエントリです。Playwright の **Browser** に相当します。

通常は `ShSetting01_StartBrowser.StartCDPMode` で取得します。タブ操作だけなら [`CDPContext`](./CDPContext) を返す `StartCDPModeContext` の方が簡単です。

```vb
Dim b As CDPBrowser
Set b = ShSetting01_StartBrowser.StartCDPMode
Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
b.quit
```

## 起動・再接続・終了

### `start`

```vb
Public Sub start(Optional Name As BrowserList = BrowserList.RunChrome, ...)
```

ブラウザを起動しパイプ接続します。日常利用では設定シート経由を推奨。

### `reattachPipe` / `reattachWebSocket` / `reattachWebView2`

Excel テーブルにある既存の接続情報（パイプ／WebSocket／WebView2）を利用して、再接続を試みます。トランスポートごとに専用のメソッドが分かれています。

```vb
Public Sub reattachPipe(userProfile As String)
Public Sub reattachWebSocket(userProfile As String, WebSocketMode As CDPCoreViaWebSocket)
Public Sub reattachWebView2(userProfile As String, WebView2Mode As CDPCoreViaWebView2)
```

| 引数 | 意味 |
| --- | --- |
| `userProfile` | 再アタッチしたいユーザー名（`user-data-dir` に基づく識別名称） |
| `WebSocketMode` | WebSocket で CDP 制御する場合、接続処理済みの `CDPCoreViaWebSocket` を指定 |
| `WebView2Mode` | WebView2 で CDP 制御する場合、接続処理済みの `CDPCoreViaWebView2` を指定 |

#### 基本的な使い方

- `start` や設定シート経由で起動したときの識別名称を `userProfile` に渡します（例: `ShSetting01_StartBrowser.CurrentUserName`）
- WebSocket / WebView2 モードで動かすときは、接続処理済みの Class オブジェクトを第 2 引数に渡します

```vb
' Pipe 版
Dim b As New CDPBrowser
b.reattachPipe ShSetting01_StartBrowser.CurrentUserName

' WebSocket 版
Dim ws As New CDPCoreViaWebSocket
ws.AutoConnectBrowserCDP UserName
Dim b As New CDPBrowser
b.reattachWebSocket UserName, ws
```

::: tip 注意
`CDPBrowser` 側はいずれも戻り値なしの `Sub` です。パイプ（または WebSocket / WebView2 接続）が生きていない場合は、`Boolean` を返さず **VBA エラーで停止**します（`CDPContext` 側の同名メソッドは `Function ... As Boolean` で、失敗時は例外を投げずに `False` を返す点が異なります）。エラーで止めたくない場合は、呼び出し側で `On Error` を使ってください。生存していない場合、Part1 からやり直す必要がある点は共通です。
:::

詳細は [再接続ガイド](/guides/reattach) / [WebSocket モード](/websocket/capabilities) を参照。

### `quit`

```vb
Public Sub quit()
```

ブラウザを終了しリソースを解放します。

### `isLiveBrowser`

```vb
Public Function isLiveBrowser() As Boolean
```

プロセス／接続が生きているか。

## タブ

### `newTab`

新規タブ（またはウィンドウ）を開き [`CDPContext`](./CDPContext) を返します。引数をすべて省略すると、空のタブをアクティブとして開きます。

```vb
Public Function newTab( _
    Optional Url As String, _
    Optional newWindow As Boolean, _
    Optional setMain As Boolean, _
    Optional isHidden As Boolean, _
    Optional browserContextId As String, _
    Optional isBackground As Boolean _
) As CDPContext
```

#### 基本的な使い方

| 引数 | 意味 |
| --- | --- |
| `Url` | 渡した URL で新規タブを開く |
| `newWindow` | `True` でタブではなく新しいウィンドウとして開く |
| `setMain` | `True` で Excel テーブルにメインタブとして記録する |
| `isHidden` | `True` で非表示タブとして開く。ブラウザ上では開いていないように見えるが、プログラム上は通常どおりタブ操作が可能 |
| `browserContextId` | 事前に `Target.createBrowserContext` で得た ID を渡すと、新しいウィンドウかつシークレット相当のコンテキストで稼働する |
| `isBackground` | `True` でタブを生成するが、アクティブにはしない |

```vb
Dim b As CDPBrowser
Set b = ShSetting01_StartBrowser.StartCDPMode

' 空タブ（アクティブ）
Dim blank As CDPContext
Set blank = b.newTab

' URL 指定 + メインタブ記録
Dim t As CDPContext
Set t = b.newTab("https://example.com", setMain:=True)
```

#### `isHidden` の使い道

例えば ZIP 解凍。普通は PowerShell 等で行いますが、どうせ Chromium が開いているならついでに JavaScript で解凍させる、という場面があります。タブは見せたくないときに重宝します。

```vb
Dim hidden As CDPContext
Set hidden = b.newTab(isHidden:=True)
' ここで JS による解凍など、UI に出したくない処理
```

#### `browserContextId` の使い道

同じログインが必要な URL に対して、異なる複数アカウントでの共通処理をしたいときに重宝します（例: 同じメニュー項目に対する PDF ダウンロードなど）。

```vb
' 事前に Target.createBrowserContext で browserContextId を取得しておく
Dim secret As CDPContext
Set secret = b.newTab( _
    Url:="https://example.com", _
    newWindow:=True, _
    browserContextId:=ctxId)
```

関連: [マルチタブ](/guides/multi-tab)

### `getTab`

既に開いているタブを、タブ名または URL に基づいて検索し、[`CDPContext`](./CDPContext) として返します。

```vb
Public Function getTab( _
    Optional tabName As String, _
    Optional Url As String, _
    Optional setMain As Boolean, _
    Optional SearchTypeID As DevToolsAgentHost_KType = kPage, _
    Optional doRetrySecond As Double _
) As CDPContext
```

| 引数 | 意味 |
| --- | --- |
| `tabName` | タブ名（第一優先・部分一致可） |
| `Url` | URL（第二優先・部分一致可） |
| `setMain` | `True` で Excel のタブ情報テーブルに上書き。タブの reattach に利用 |
| `SearchTypeID` | `kPage` / `kFrame` など、ターゲット種類。既定は `kPage` |
| `doRetrySecond` | 指定秒以内に見つかるまでリトライ。`0`（既定）なら 1 回のみ |

```vb
Dim b As CDPBrowser
Set b = ShSetting01_StartBrowser.StartCDPMode

' 直近の未接続 page タブへ（tabName / Url を省略）
Dim t As CDPContext
Set t = b.getTab(setMain:=True)

' タイトル／URL で検索
Set t = b.getTab(tabName:="Example", Url:="https://example.com")

' iframe を最大 5 秒リトライして探す
Set t = b.getTab( _
    Url:="https://challenges.cloudflare.com/", _
    SearchTypeID:=kFrame, _
    doRetrySecond:=5)
```

::: tip 注意
- `tabName` と `Url` を両方省略すると、最も近い未接続のタブに接続しようとします
- reattach 後は `setMain:=True` を推奨します
:::

#### `SearchTypeID`（`DevToolsAgentHost_KType`）

Chromium の [DevToolsAgentHost 種別](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/devtools_agent_host_impl.cc?ss=chromium&q=f:devtools%20-f:out%20%22::kTypeTab%5B%5D%22) に準拠した Enum です。種類は多いですが、覚えるのは次の 2 つで十分です。省略時は `kPage` なので、ほとんどのケースでは意識不要です。

| 値 | 意味 |
| --- | --- |
| `kPage` | 普段目にするタブ（既定） |
| `kFrame` | iframe と同等 |

その他（必要になったとき）: `kTab` / `kDedicatedWorker` / `kSharedWorker` / `kServiceWorker` / `kWorklet` / `kBrowser` / `kGuest`（webview） / `kOther`（非表示タブ） / `kAuctionWorklet` / `kAssistiveTechnology` / `kBrowserUI`

関連: [マルチタブ](/guides/multi-tab) / [再接続](/guides/reattach)

### `PageCount`

```vb
Public Function PageCount() As Long
```

`Target.getTargets` の結果のうち、`"type" = "page"` のターゲットだけを数えます（普段目にするタブ相当）。

### `attachToTab` / `DiscardSessionID`

```vb
Public Function attachToTab(tabId As String) As String
Public Function DiscardSessionID(sessionID As String) As Boolean
```

[`CDPContext`](./CDPContext) ↔ `CDPBrowser` のやり取り用として公開している低レベル API です。日常利用では意識不要です。

自前でタブ管理したいときの **タブ接続（`attachToTab`）／セッション破棄（`DiscardSessionID`）** として使えます。

## プロトコル

### `ExecuteCDP` / `ExecuteCDPAsync`

```vb
Public Function ExecuteCDP(methodName As String, _
    Optional params As Dictionary, _
    Optional StopCDPError As Boolean = True) As BiDiCDPJson

Public Function ExecuteCDPAsync(...) As Long
```

ブラウザターゲット向け CDP コマンドです。結果を待つか（`ExecuteCDP`）、待たずに後で確認するか（`ExecuteCDPAsync`）の 2 種類を用意しています。

`ExecuteCDPAsync` はコマンド実行時の **id（`Long`）のみ**を返し、結果は待ちません。結果の回収は [`TakeResultCDP`](#takeresultcdp) で行います。

詳細は [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol) で説明します。

### `TakeEvents`

```vb
Public Sub TakeEvents(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
```

非同期応答／イベントの吸い上げ（**almighty**：受信 → 解析 → 蓄積中の全メッセージについて `RaiseEvent` までを一括で行う）。`TakeResultCDP` の前に呼び出す必要があります。

| 引数 | 意味 |
| --- | --- |
| `StopApiError` | Pipe／WebSocket 障害時に停止するか。既定は `True` |
| `destruction` | `True` でストリームに蓄積せず破棄するだけ（`RaiseEvent`は起きない） |

### `TakeEvent` / `NewResCDP` / `AnalyzeCDP`

```vb
Public Sub AnalyzeCDP(Optional StopApiError As Boolean = True, Optional destruction As Boolean)
Property Get NewResCDP() As Boolean
Public Sub TakeEvent()
```

`TakeEvents` を分解した **manual** な低レベル部品です。特定イベント検出時に即座に割り込んで残りのイベント処理を止めたい、といった上級者向け用途で使います。

| メンバー | 役割 |
| --- | --- |
| `AnalyzeCDP` | Pipe／WebSocket からバイト配列を受け取り、テキストへ変換して 1 件分取り出せる状態にする |
| `NewResCDP` | `AnalyzeCDP` 後、取り出せる新規メッセージがあるか（`Get` 専用） |
| `TakeEvent` | 蓄積済みメッセージを **1 件だけ**取り出し `RaiseEvent` する |

```vb
' TakeEvents は、概ね次と同義
Do
    b.AnalyzeCDP
    Do While b.NewResCDP
        b.TakeEvent
    Loop
Loop While ...
```

### `TakeResultCDP`

```vb
Property Get TakeResultCDP(commandID As Long) As String
```

`ExecuteCDPAsync` が返した id をキーに、蓄積された実行結果（JSON 文字列）を取り出します。取り出し後は Dictionary から削除されます。結果がまだ無い場合は空文字を返します。

### `SetLimitCDPResult`

```vb
Property Let SetLimitCDPResult(Number As Long)
```

CDP コマンド結果を Dictionary に溜め込む件数の上限です。デフォルトは **65536 件**です。

上限を超えると、パフォーマンス低下を防ぐため蓄積中の結果履歴が **すべて削除**されます。未回収の `ExecuteCDPAsync` 結果も消える点に注意してください。

```vb
b.SetLimitCDPResult = 1000
```

::: tip
コマンド ID がオーバーフロー対策でリセットされるとき（およそ 20 億到達時）も、結果履歴はすべてクリアされます。
:::

### `LastCDPJsonError`

直前エラー（Dictionary 風アクセス）。`StopCDPError:=False` 時に参照。

## タイムアウト

### `TimeOutSecond`

```vb
Property Let TimeOutSecond(TimeSec As Double)
```

CDP コマンド結果待ちなどの待機上限です。デフォルトは **30 秒**です。**LET 専用**（書き込みのみ）で、設定中の値は読み返せません。

```vb
b.TimeOutSecond = 60
```

詳細は [タイムアウト設定方法について](/guides/timeout)。

## その他

| メンバー | 説明 |
| --- | --- |
| `openDevTools` | 指定ターゲットで DevTools を開く |
| `printTargetInfos` / `printParams` | デバッグ出力 |
| `sleep` | 秒待ち |
| `TimerCounter` | 経過ミリ秒。`Timer` 関数の代わりに自前ループのタイムアウト判定へ。[タイムアウト設定方法について](/guides/timeout) |
| `serializeForMainTab` | メインタブの session/target を記録 |

## 関連

- [マルチタブ](/guides/multi-tab)
- [タイムアウト設定方法について](/guides/timeout)
- [`CDPContext`](./CDPContext)
