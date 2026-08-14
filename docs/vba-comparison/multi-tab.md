---
description: 1本の接続で複数タブをどう捌くか。sessionId 多重化・グローバル現在地・都度再接続という3つの実装と、非同期実行の可否、そして接続エンドポイントの選び方を比較します。
---

# マルチタブとセッション管理

CDP には「ウィンドウハンドル」という概念がありません。複数タブを扱いたければ、`Target` ドメインと `sessionId` を使って**各ツールが独自にタブ管理の仕組みを発明する**必要があります。3プロジェクトで実装の充実度に大きな差が出たのは、そのためです。

## まず入口 ―― どのエンドポイントで繋ぐか

WebSocket でブラウザに繋ぐとき、Chromium は2つの入口を用意しています。

| エンドポイント | 得られるもの | その後 |
| --- | --- | --- |
| `/json/version` | **ブラウザ全体**の `webSocketDebuggerUrl` | `Target.attachToTarget` + `sessionId` を自力で組み立てる |
| `/json/list`（`/json`） | **ページごと**の `webSocketDebuggerUrl` | そのページ専用ソケットに直接繋ぐだけ |

前者は面倒な代わりにブラウザ全体を掌握でき、後者は楽な代わりに1ページに固定されます。3者の選択はこう分かれました。

- **VBAChromeDevProtocol** は `/json/version` 一本槍
- **vba-cdp-webdriver** は `/json/list` 一本槍
- **StarterWebScrapingKit** は `AutoConnectPageCDP`（`/json/list` 経由）と `AutoConnectBrowserCDP`（`/json/version` 経由）の**両方を用意**し、用途で選べるようにしています

```vb
'「/json/list」へアクセスして、利用可能なすべてのウェブソケットターゲットのリストのうち、
' 引数に基づいた`Page`接続のWebSocket接続まで行います
Public Function AutoConnectPageCDP(UserName As String, Optional Url As String, Optional Title As String, ...) As Boolean

'「/json/version」へアクセスして、ブラウザ接続のWebSocket接続まで行います
Public Function AutoConnectBrowserCDP(UserName As String, Optional ReuseContext As Boolean, ...) As Boolean
```

「もう分かっているページにサクッと繋ぎたい」ときは楽な道、「ブラウザ全体を掌握して後から好きなタブを発見・アタッチしたい」ときは大変な道。**どちらか一方に絞らないこと自体が、継続投資の一形態**です。

## タブの実体 ―― 3つの答え

### StarterWebScrapingKit：タブごとに独立したオブジェクト

`newTab()` / `getTab()` は、それぞれ独立した `CDPContext` オブジェクトを返します。各 `CDPContext` は自分専用の `CurrentSessionID` を持ち、送信時に自分のセッション ID をメッセージへ載せます。接続（Pipe / WebSocket）は1本を全タブで共有・多重化します。

```vb
tab1.ExecuteCDP "Page.navigate", params1
tab2.ExecuteCDP "Page.navigate", params2   ' ← 「切り替え」という手順が存在しない
```

グローバルな「現在のタブ」ポインタを持たないため、複数タブを行き来するコードでも状態管理のバグが起きにくい構造です。`getTab` はタイトル / URL の部分一致にスコアリングとリトライ待機まで実装されていて、似た名前のタブが並んでいても狙ったものを掴めます。

### VBAChromeDevProtocol：多重化はするが「現在地」はグローバル1個

同一パイプ上で `sessionId` 多重化を行うので、**タブ切り替えに再接続は不要**です。ここは同じ答えに辿り着いています。

ただし `sessionId` は `clsCDP` インスタンスにつき1つの、グローバルな「現在地」ポインタです。

```vb
' current session attached to direct message to
' Note: sendMessage automatically adds to each message sent to browser
'       set to vbNullString to send avoid adding sessionId to message (sends to browser sessionless target)
Public sessionId As String
```

`switchTo` はこのプロパティを書き換えるだけなので切り替え自体は高速ですが、**呼び忘れると意図しないタブにコマンドが飛びます**。「タブAとタブBを交互に操作する」コードでは、都度この呼び出しに気をつける必要があります。

### vba-cdp-webdriver：切り替えのたびにソケットを張り直す

表面上の API は3つの中で最も親切です。`SwitchTabByIndex` / `SwitchTabByTitle` / `GetTabList` / `CloseTabByTitle` と、Selenium 作法そのままの直感的な名前が並びます。

しかし中身は違いました。

```vb
Private Sub ReconnectToTarget(ByVal targetId As String)
    Dim tabs As Collection: Set tabs = GetPageTabs()          ' ① http://127.0.0.1:9222/json へ HTTP リクエスト
    webSocketURLPath = FindWebSocketPathByTargetId(tabs, targetId)

    If Not webSocket_ Is Nothing Then webSocket_.CloseWebSocket ' ② 既存ソケットを切断

    Set webSocket_ = New a1_WebSocketCommunicator              ' ③ 新規ソケットを張り直す
    If webSocket_.Init(webSocketURLPath) = False Then ...

    Set CDP_ = New a4_ExecuteCDP                               ' ④ 実行層オブジェクトも作り直す
    CDP_.Init msgGenerator, Handler, json_, events_

    CDP_.AttachToTarget targetId                               ' ⑤ ここから CDP 5往復
    CDP_.ActivateTarget targetId
    CDP_.PageEnable
    CDP_.DOMEnable
    CDP_.RuntimeEnable
```

タブを1回切り替えるたびに、**HTTP リクエスト1回 + ソケット張り直し + CDP コマンド5往復**が走ります。どの瞬間を切り取っても生きている接続は1本だけで、同時並行の操作はできません。2つのタブを細かく行き来する処理では、このオーバーヘッドが積み上がります。

多重化への未練は残っています。`Target.attachToTarget` の応答から `sessionId` を取り出して保持する処理はあるのですが。

```vb
SessionId = json_.GetValue(res, "result", "sessionId")
MsgGenerate_.SessionId = SessionId
```

このプロパティを**読んでいる場所がどこにもありません**。送信 JSON に `sessionId` を載せるコードは存在せず、`Target.attachToTarget` にも `flatten` を渡していません。ページ専用ソケットに直接繋いでいる以上どのみち不要なので、多重化に踏み出す手前で止まった痕跡がそのまま残っている、という状態です。

## 「クリックしたら別タブが開いた」への対応

よくある場面ですが、対応の方向性が分かれています。

**VBAChromeDevProtocol** は `attachToChildTargets` を持っていて、`openerId`（どのターゲットから開かれたか）を見て子ウィンドウを検出・アタッチします。ただし自動で追跡してくれるわけではなく、クリック時に「新しいウィンドウが開くかも」と明示的に宣言しておく必要があります。

```vb
browser.Click node.nodeId, strategy:=WindowOpen
```

親子関係を `openerId` で正確に辿るので、**同時に複数のタブが開いても取り違えません**。ここは3者で最も厳密です。

**StarterWebScrapingKit** は `getTab` のタイトル / URL 一致 + スコアリング + リトライで拾いにいきます。宣言が不要で書きやすい代わりに、親子関係そのものを見ているわけではありません。

**vba-cdp-webdriver** は `GetTabList` を取り直して `SwitchTabByTitle` する、という手動の流れになります。

## 非同期実行 ―― 結果を待たずに撃てるか

CDP の高速化テクニックのひとつが「独立したコマンドをまとめて撃って、後から結果を回収する」ことです。ここははっきり差が出ました。

**StarterWebScrapingKit** は公開 API として実装済みです。

```vb
' コマンド実行IDだけをもらって即座に処理を返す（結果を待たない）
id1 = tab.ExecuteCDPAsync("Runtime.evaluate", params1)
id2 = tab.ExecuteCDPAsync("Runtime.evaluate", params2)
id3 = tab.ExecuteCDPAsync("Runtime.evaluate", params3)
' ↑ 3回分のラウンドトリップ待ちをせず、ほぼノータイムで発射しきれる

result1 = tab.TakeResultCDP(id1)   ' 届いていなければ空文字。整理券番号で後から回収
```

**VBAChromeDevProtocol** は土台だけ持っています。`sendMessage` には `nowait` 引数があるのですが。

```vb
' but if called with nowait=True then returns without waiting for reply - you must call peakMessage in a loop
Private Function sendMessage(ByVal strMessage As String, Optional ByVal nowait As Boolean = False) As Dictionary
```

`Private` です。公開 API の `InvokeMethod` はこれを渡しておらず、自動生成された `cdp.Page.navigate` などのドメインクラスも**すべて** `InvokeMethod` 経由なので、ライブラリが提供する通常の呼び出し経路は 100% 同期・ブロッキングになります。`peakMessage` が `Public` なので理論上は自作できますが、examples にも前例はありません。

**vba-cdp-webdriver** には仕組み自体がありません。

### 非同期の前提になる「覗き見」

「結果を待たずに撃つ」ためには、受信側が**データが届いているかどうかを、ブロックせずに確認できる**必要があります。ここが VBA では最大の関門でした。

vba-cdp-webdriver の受信ループは `WinHttpWebSocketReceive` を直接叩いています。

```vb
Do
    res.result = WinHttpWebSocketReceive( _
                    http.websockethandle, res.Buffer(0), bufLen, _
                    res.ReceiveBytes, res.status)
    ' ... 最終フラグメントまで集めて復元
Loop
```

この API には peek 機能がないので、何も届いていないときに呼ぶと**届くまで戻ってきません**。VBAChromeDevProtocol も同じ API を使っていて、作者自身がその副作用をコメントに残しています。

```vb
dwError = WinHttpWebSocketReceive(hWebSocketHandle, rgbBuffer(dwBytesResidue), ...)
DoEvents ' the WinHttpWebSocketReceive call will make the VBA Application unresponsive until it returns, give slice back
```

「戻ってくるまで Excel は応答不能になる」と明記されています。`DoEvents` は戻ってきた**後**にしか効かないので、根本的な解決にはなっていません。

面白いのは、VBAChromeDevProtocol の**本流である Pipe 側にはちゃんと覗き見がある**ことです。

```vb
' Reading is non-blocking, if there are no bytes to read the function returns 0
Public Function readProcCDP(ByRef strData As String) As Long
    Call PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, lPeekData, ByVal 0&)
    If lPeekData > 0 Then
        ' ...
    Else
        readProcCDP = -1   ' データなし
    End If
```

`PeekNamedPipe` が使える Pipe では3者とも非ブロッキングに辿り着いていて、**WebSocket に移った途端に道が消える**。つまりこれは実装の巧拙ではなく、選んだ API に peek があるかどうかの問題でした。

StarterWebScrapingKit の `CDPCoreViaWebSocket.cls` は、この2択を避けるために WinSock まで降りて RFC 6455 を自前実装するという道を選びました。ヘッダコメントにその判断が書かれています。

> **なんで、「WinHttpWebSocket○○」系の API を使わないの？**
> 「WinHttpWebSocket○○」の場合、`ioctlsocket` や `PeekNamedPipe` と言った「覗き見機能」はなく、下記の2択しか選択肢がありません
> ・OS の推奨する「非同期コールバック」… VBE で一時停止したり、デバッグでマクロを途中で止めたりした瞬間に、Excel ごと警告なしでクラッシュ
> ・同期モード … Peek がないから、データがまだ届いていないときに1行呼んだ瞬間、Excel が「応答なし」で完全にフリーズ

`ioctlsocket(FIONREAD)` で「今何バイト来ているか」を先に聞ければ、届いていないときは素通りできます。**非同期実行 API が3者で1つしか成立しなかった理由は、突き詰めるとこの1点**でした。

## 補足：SeleniumVBA との対比

「WebDriver」という名前が付いた VBA プロジェクトはもう1つあります。[SeleniumVBA](https://github.com/GCuser99/SeleniumVBA) です。名前は vba-cdp-webdriver に似ていますが、マルチタブの充実度は**真逆**でした。

```vb
driver.Windows.SwitchToByTitle("New Window")
driver.Windows.SwitchToNew(windowType:=svbaTab)
For Each window In driver.Windows
```

`WebWindows.cls` という専用のコレクションクラスを持ち、最初から複数ウィンドウ前提です。理由は明快で、SeleniumVBA は **CDP を直接叩いていません**。

```vb
.CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
.CMD_NEW_WINDOW = Array("POST", "/session/$sessionId/window/new")
```

W3C WebDriver 仕様の REST API そのままです。**「ウィンドウハンドルの一覧を取得する」という概念が仕様に標準機能として定義されている**ため、ラップするだけで済みます（CDP は `ExecuteCDP` という抜け道として別途用意されているだけ）。

CDP 直叩き勢がマルチタブで苦労したのは、実装が拙かったからではなく、**そもそも標準機能が存在しない土俵を選んだから**です。名前が似ている2つのプロジェクトが対照的な結果になったのは、この選択の違いでした。

## まとめ

| | タブ切替の実体 | 同時並行 | 非同期実行 |
| --- | --- | --- | --- |
| **StarterWebScrapingKit** | `sessionId` 多重化（再接続なし） | タブオブジェクトごとに独立、切替不要 | `ExecuteCDPAsync` + `TakeResultCDP` |
| **VBAChromeDevProtocol** | `sessionId` 多重化（再接続なし） | グローバル「現在地」を都度切替 | `nowait` は Private、実質使えない |
| **vba-cdp-webdriver** | ソケット都度張り直し + CDP 5往復 | 排他的、1タブずつ | なし |

「何十個もタブを同時に触りたい」なら StarterWebScrapingKit が頭一つ抜けています。「ポップアップの親子関係を正確に追跡したい」というピンポイントな用途なら VBAChromeDevProtocol の `openerId` ベースの追跡に分があります。

## 関連

- [非同期実行と待機戦略](/core-comparison/async) — Node.js の Promise / イベントループとの比較
- [トランスポート層とバッファ管理](/core-comparison/transport) — WinSock 自前実装の詳細
- [WebSocket 設計](/websocket/design) — なぜ WebSocket 対応を後から足したのか
