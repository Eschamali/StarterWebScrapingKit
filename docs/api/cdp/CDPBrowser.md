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
b.CloseBrowser
```

## 起動・再接続・終了

### `OpenBrowser`

```vb
Public Function OpenBrowser(caps As BrowserCapabilities) As String
```

ブラウザを起動し、Pipe / Port（WebSocket）いずれかで接続します。どちらで接続するかは、`caps` の `Set_UseWebSocket`（既定はワークシートの「WebSocketモード」。OFF なら Pipe）で決まります。戻り値は初期接続先の URL 文字列です。

| 引数 | 意味 |
| --- | --- |
| `caps` | 起動時のオプションをまとめた [`BrowserCapabilities`](./BrowserCapabilities) |

```vb
Dim b As CDPBrowser
Set b = New CDPBrowser
b.OpenBrowser BrowserCapabilities
```

日常利用では設定シート経由（`StartCDPMode` 等）を推奨します（内部でこのメソッドを呼んでいます）。

### `reattach`

Excel テーブルにある既存の接続情報を利用して、起動済みのブラウザへ再接続します。通信経路（Pipe / Port / WebView2）は、[`ReattachOptions`](./ReattachOptions) で指定します。

```vb
Public Sub reattach(reattachInfo As ReattachOptions)
```

| 引数 | 意味 |
| --- | --- |
| `reattachInfo` | 再接続時のオプションをまとめた [`ReattachOptions`](./ReattachOptions)。制御経路の指定が必須 |

```vb
' Pipe 版
ReattachOptions.Set_PipeMode = ShSetting01_StartBrowser.CurrentUserName

Dim b As New CDPBrowser
b.reattach ReattachOptions

' Port（WebSocket）版
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket
ws.AutoConnectBrowserCDP UserName

ReattachOptions.Set_WebSocketMode(ws) = UserName

Set b = New CDPBrowser
b.reattach ReattachOptions
```

::: tip 注意
戻り値なしの `Sub` です。パイプ（または WebSocket / WebView2 接続）が生きていない場合は、`Boolean` を返さず **VBA エラーで停止**します（`CDPContext.reattach` は `Function ... As Boolean` で、失敗時は例外を投げずに `False` を返す点が異なります）。エラーで止めたくない場合は、呼び出し側で `On Error` を使ってください。生存していない場合、Part1 からやり直す必要がある点は共通です。
:::

詳細は [再接続ガイド](/guides/reattach) / [WebSocket モード](/websocket/capabilities) を参照。

### `CloseBrowser`

```vb
Public Sub CloseBrowser()
```

`Browser.close` を送ってブラウザを終了し、Excel テーブル上のメインタブ情報と接続ハンドルを解放します。一定時間（5 秒）たっても閉じない場合は、プロセスを強制終了します。

```vb
b.CloseBrowser
```

### `DisconnectCDP`

```vb
Public Sub DisconnectCDP()
```

**ブラウザは終了せずに**、接続だけを切り離し、Excel テーブルの接続情報も削除します。

::: warning Port（WebSocket）以外では再接続できなくなります
Pipe / WebView2 で `DisconnectCDP` を呼ぶと、そのブラウザへは二度と再接続できなくなります。ブラウザを開いたまま、あとで [`reattach`](#reattach) したい場合は、`DisconnectCDP` を呼ばずに処理を終えてください（Port の場合のみ、切断したあとでも再接続できます）。
:::

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

::: warning v3.2.0での変更
`Url` を指定した `newTab` は、以前は読み込み完了までブロッキングしていましたが、v3.2.0以降は**即座に戻る**ようになりました（内部の待機処理が `Wait`（ポーリング）から `PageReadyState`（現在値の取得のみ）に変わったため）。読み込み完了を待ちたい場合は、戻り値の `CDPContext` に対して明示的に `.Wait` を呼んでください。
```vb
Dim t As CDPContext
Set t = b.newTab("https://example.com", setMain:=True)
t.Wait   ' 読み込み完了まで待つ（v3.2.0以降は明示的に必要）
```
:::

#### `isHidden` の使い道

例えば ZIP 解凍。普通は PowerShell 等で行いますが、どうせ Chromium が開いているならついでに JavaScript で解凍させる、という裏処理場面があります。内部処理部分のタブは見せたくないときに重宝します。

```vb
Dim hidden As CDPContext
Set hidden = b.newTab(isHidden:=True)
' ここで JS による解凍など、UI に出したくない処理
```

#### `browserContextId` の使い道

ログインが必要な、しかし**同一URL**に対して、異なる複数アカウントでの共通処理をしたいときに重宝します（例: 同じメニュー項目に対する PDF ダウンロードなど）。

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
- 初回reattach 後は `setMain:=True` を推奨します
:::

#### `SearchTypeID`（`DevToolsAgentHost_KType`）

`DevToolsAgentHost_KType` は、`CDPHelpers` モジュールで定義している Enum です。Chromium の [DevToolsAgentHost 種別](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/devtools_agent_host_impl.cc?ss=chromium&q=f:devtools%20-f:out%20%22::kTypeTab%5B%5D%22) に準拠した Enum です。種類は多いですが、覚えるのは次の 2 つで十分です。省略時は `kPage` なので、ほとんどのケースでは意識不要です。

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
Public Sub DiscardSessionID(sessionID As String)
```

[`CDPContext`](./CDPContext) ↔ `CDPBrowser` のやり取り用として公開している低レベル API です。日常利用では意識不要です。

自前でタブ管理したいときの **タブ接続（`attachToTab`）／セッション破棄（`DiscardSessionID`）** として使えます。

### `getTabForOpenerID`

```vb
Public Function getTabForOpenerID( _
    PageID As String, _
    Optional Strict As Boolean, _
    Optional doRetrySecond As Double _
) As CDPContext
```

指定したタブ（または iframe）から開かれたタブ・ウィンドウ（主にポップアップ）を探し、[`CDPContext`](./CDPContext) として返します。条件に合うタブがない場合は `Nothing` です。タブ側から呼ぶ [`CDPContext.PopupCatch`](./CDPContext#popupcatch) は、このメソッドに自分の `CurrentTargetID` を渡しています（自分が iframe のターゲットなら `Strict:=True`）。

| 引数 | 意味 |
| --- | --- |
| `PageID` | 開いた元の `TargetID`。`Strict:=True` のときは `FrameId` |
| `Strict` | `False`（既定）で `PageID` を `TargetID` として、`True` で `FrameId` として扱う |
| `doRetrySecond` | 何秒間探すか。`0`（既定）なら 1 回のみ |

使うときは、次の手順を踏みます。

1. `Set b.RunCollect_TargetCreated = New Collection` で、タブが作られたこと（`Target.targetCreated`）の記録を始める
2. 新しいタブ・ウィンドウが開く操作（クリックなど）をする
3. 開いた元の `CurrentTargetID` を `PageID` に渡す。iframe の中で操作した場合は、`Strict:=True` にして `FrameId` を渡す
4. クリックの直後に `Target.targetCreated` が届かないことがあるので、必要に応じて `doRetrySecond` で待つ時間を調整する
5. 条件に合うタブが見つかると、それを `CDPContext` として返し、記録からも取り除く
6. 用が済んだら、`Set b.RunCollect_TargetCreated = Nothing` で記録を終わらせる

```vb
Set b.RunCollect_TargetCreated = New Collection
t.getElementByQuery("a[target=_blank]").click

Dim popup As CDPContext
Set popup = b.getTabForOpenerID(t.CurrentTargetID, doRetrySecond:=5)

Set b.RunCollect_TargetCreated = Nothing
```

::: tip 待機と組み合わせるとき
見つけたタブのページの読み込みを待つ場合は、[`CDPContext.Wait`](./CDPContext#wait)（JavaScript による待機）を使ってください。[`CDPContext.WaitEvents`](./CDPContext#waitevents) は、イベントの届く順番の都合で、ここでは機能しません。
:::

## プロトコル

### `ExecuteCDP` / `ExecuteCDPAsync`

```vb
Public Function ExecuteCDP(methodName As String, _
    Optional params As Dictionary, _
    Optional StopCDPError As Boolean = True) As BiDiCDPJConv

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

::: tip
コマンド回数が20億を超えると、オーバーフロー対策により蓄積中の結果履歴が **すべて削除**されます。未回収の `ExecuteCDPAsync` 結果も消える点に注意してください。
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
| `printTargetInfos` / `printParams` | デバッグ出力。`printParams` は、`Browser.getVersion` から得たブラウザ情報（製品名・バージョン・V8 / CDP バージョン・リビジョン・`User-Agent`）と、ユーザーデータ名・プロセスID・ツールのバージョンを表示します |
| `CurrentBrowserUser`（Get） | 起動中ブラウザのユーザーデータ識別名称 |
| `CurrentBrowseruserAgent`（Get） | 起動中ブラウザの`User-Agent`値（接続時に `Browser.getVersion` から取得してキャッシュしたもの） |
| `UseCustomUAMode`（Get） | このツールのサインを `User-Agent` に追記する設定が有効か（起動時は `BrowserCapabilities.Set_EnableUserAgentOverride`、再接続時は `ReattachOptions.Set_EnableUserAgentOverride` で決まる） |
| `ExEvent`（イベント） | `Event ExEvent(methodName As String, RawJson As String)`。自作の拡張クラス向けに、ブラウザ単位の CDP 非同期イベント（`Target.targetCreated` など）を通知します。`Private WithEvents ex_CDPBrowser As CDPBrowser` で受け取ります。タブ宛てのイベントは [`CDPContext`](./CDPContext) の `ExEvent` です |
| `sleep` | 秒待ち |
| `TimerCounter` | 経過ミリ秒。`Timer` 関数の代わりに自前ループのタイムアウト判定へ。[タイムアウト設定方法について](/guides/timeout) |
| `serializeForMainTab` | メインタブの session/target を記録 |
| `RunCollect_TargetCreated`（Set） | `Target.targetCreated` の記録の開始（`New Collection`）と終了（`Nothing`）。[`getTabForOpenerID`](#gettabforopenerid) の前後で使います |

## 関連

- [マルチタブ](/guides/multi-tab)
- [タイムアウト設定方法について](/guides/timeout)
- [`CDPContext`](./CDPContext)
- [`BrowserCapabilities`](./BrowserCapabilities) / [`ReattachOptions`](./ReattachOptions)
