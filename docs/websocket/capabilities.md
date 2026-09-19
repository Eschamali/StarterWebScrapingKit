---
description: WebSocket モードでできること。Android ブラウザや WebView2 など、Pipe では届かない遠隔・埋め込み制御の使いどころを紹介します。
---

# WebSocket モードでできること

ほとんどのブラウザ自動化は Pipe で十分ですが、以下のシチュエーションで自動化する場合は WebSocket モードを使うことになります。

- [Android ブラウザ制御](https://developer.chrome.com/docs/devtools/remote-debugging?hl=ja)
![Androidブラウザが制御されてる様子](../public/viaWebSocket/Android制御.gif)

- [WebView2 制御](https://playwright.dokyumento.jp/docs/webview2)（**他アプリ**に組み込まれた WebView2 を、デバッグポート越しに後付け制御する場合）
![WebView2 制御してる様子](../public/viaWebSocket/WebView2制御.avif)

- [今目の前のブラウザを制御](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja)
![今目の前のブラウザを制御してる様子](/viaWebSocket/目の前のブラウザ制御.gif)

- Tailscale 等によるインターネットを介した制御

::: tip Excel自身に埋め込むWebView2の場合
Excel（VBA）自身のUserFormにWebView2を埋め込んで制御したいだけなら、WebSocketモードではなく [WebView2モードでの制御について](/webview2/design) というネイティブな専用トランスポート（v3.0.0〜）を使ってください。デバッグポートもWebSocketも経由しない、より直接的な経路です。
:::

## ローカルブラウザの起動から行う場合

ローカルブラウザの**起動から接続まで**を一気に行う方法が2通りあります（v3.0.0〜、v3.1.0でAPI整理）。

### 設定シート経由（もっとも簡単）

ブラウザ起動設定シートの **`UseWebSocket` セルを `TRUE`** にするだけで、Pipe の代わりに WebSocket 経由でローカルブラウザを起動・接続するようになります（既定は `FALSE` = Pipe）。

![ブラウザ起動設定シートの基本設定欄。「WebSocketモード」というチェックボックス行が赤枠で強調されており、右側に「新規起動時の制御経路を設定します。OFFで...」という説明が添えられている](/img/セルからWebSocket切替.png)

*▲ シート上では「WebSocketモード」という表記のチェックボックスです（`UseWebSocket`はVBA側の内部名）*

```vb
Sub AutoConnectBrowser()
    '1. UseWebSocket セルが TRUE の状態で、ブラウザを起動
    Dim BrowserControl As CDPBrowser
    Set BrowserControl = ShSetting01_StartBrowser.StartCDPMode

    '2. 未接続のタブに接続
    Dim t As CDPContext
    Set t = BrowserControl.getTab(setMain:=True)

    '3. ページ遷移
    t.navigate "https://example.com"

    '4. 終了
    BrowserControl.quit
End Sub
```

::: warning v3.2.0での変更
以前は `StartCDPMode(WebSocketMode:=True)` のように、`StartCDPMode` / `StartBiDiMode` へ渡す `WebSocketMode As Boolean` 引数で切り替えていました。v3.2.0でこの引数は廃止され、ワークシートの `UseWebSocket` セルに一本化されています。あわせて、これまで非対応だった **`StartCDPModeContext` / `StartBiDiModeContext`**（`CDPContext` / `WebDriverBiDiContext` を直接返す方）でも WebSocket 起動に対応しました。つまりこの4種類の起動ヘルパー全てが、同じ `UseWebSocket` セル1つで挙動を切り替えます。
:::

### `CDPCoreViaWebSocket` を直接使う場合

内部で何が起きているかを制御したい場合は、`ConnectCDPWithLocalBrowser`（起動＋接続のみ）と、`CDPBrowser.reattachWebSocket` / `WebDriverBiDiMode.reattach`（接続済みの状態を `CDPBrowser` / `WebDriverBiDiMode` として受け取る）を組み合わせて呼べます。

```vb
Public Function ConnectCDPWithLocalBrowser(UserName As String, appUrl As String, SplashScreenMode As Boolean, addArgs As String) As String
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称（必須・第1引数） |
| `appUrl` | `--app` に付ける URL |
| `SplashScreenMode` | `True` で実URLの前に起動スプラッシュ画面を経由（詳細は [ページ遷移](/guides/navigation)） |
| `addArgs` | 追加の起動引数 |

戻り値は初期接続先の URL 文字列です。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

Dim ws As New CDPCoreViaWebSocket
ws.ConnectCDPWithLocalBrowser UserName, "https://example.com", False, vbNullString

Dim b As New CDPBrowser
b.reattachWebSocket UserName, ws

Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
```

内部では、リモートデバッグを禁止するポリシーのチェック・残存セッションの後始末・クラッシュ復元プロンプトの無効化・起動コマンドライン生成・`DevToolsActivePort` の出現待機・接続までを [`CDPHost`](/concepts/architecture) に委託したうえで自動的に行います。

::: warning v3.2.0での変更
`BrowserType As BrowserList` 引数が廃止されました（Chrome / Edge は設定シートの `UseChrome` セルで選択）。また、`CDPCoreViaWebSocket.StartCDPMode()` / `StartBiDiMode(sessionCapabilitiesRequest)`（接続済み情報を `CDPBrowser` / `WebDriverBiDiMode` に変換する専用メソッド）は**廃止**されました。依存の向きが逆転し、`CDPBrowser.start` / `WebDriverBiDiMode.StartBiDiMode` の方が内部で `CDPCoreViaWebSocket` を生成するようになったためです。上記のとおり、代わりに `reattachWebSocket` / `reattach` を組み合わせてください。日常利用では、前項の設定シート経由（`UseWebSocket` セル）の方が簡単です。
:::

## 基本的な接続方法（既存ブラウザへの後付け接続）

前節の「起動から行う場合」を除き、WebSocket は「後付け」接続のため、Pipe 版の `Start○○ModeContext` とは流れが違います。大まかには次のとおりです。

1. **接続の識別名称を取得／設定** — セル（`ShSetting01_StartBrowser.CurrentUserName`）から取ってもよいし、独自の名前でも OK
2. **目的に合った接続メソッドを呼ぶ** — 下の 3 種類から選択（`CDPCoreViaWebSocket`）
3. **対応する `reattachWebSocket` に、2. の Class オブジェクトを渡す** — Page 接続なら `CDPContext`、Browser 系なら `CDPBrowser`（または BiDi 側の Mode の `reattach`）
4. **あとはいつも通りの制御**

各 Demo モジュールの **「WebSocket経由版Demo」** セクションを参照してください（`Demo_CDP` / `Demo_WebDriverBiDi` / `Demo_WebSocket`）。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

Dim ws As New CDPCoreViaWebSocket
If Not ws.AutoConnectPageCDP(UserName) Then Exit Sub

Dim t As New CDPContext
If Not t.reattachWebSocket(UserName, ws) Then Exit Sub

t.navigate "https://example.com"
ws.DisconnectCDP
```

## 接続の種類

`CDPCoreViaWebSocket` には、既存ブラウザへ後付け接続する次の 3 種類の公開メソッドに加え、前述の起動込みメソッドがあります。

| メソッド | エンドポイント／手段 | 渡す `reattachWebSocket` |
| --- | --- | --- |
| `AutoConnectPageCDP` | `/json/list` → Page | [`CDPContext`](/api/cdp/CDPContext) |
| `AutoConnectBrowserCDP` | `/json/version` → Browser | [`CDPBrowser`](/api/cdp/CDPBrowser) |
| `ReConnectCDP`（v3.2.0〜） | `DevToolsActivePort` ファイル、または記録済みハンドルの再利用 | [`CDPBrowser`](/api/cdp/CDPBrowser) |
| `ConnectCDPWithLocalBrowser` + `reattachWebSocket`/`reattach` | ローカルブラウザを起動してから接続 | （前述の「`CDPCoreViaWebSocket` を直接使う場合」を参照） |

### `AutoConnectPageCDP`

`/json/list` へアクセスし、利用可能な WebSocket ターゲットのうち、引数に基づいた **Page** 接続まで行います。

```vb
Public Function AutoConnectPageCDP( _
    UserName As String, _
    Optional Url As String, _
    Optional Title As String, _
    Optional port As Long = 9222, _
    Optional Host As String = "127.0.0.1" _
) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称 |
| `Url` | ページ URL（部分一致など） |
| `Title` | ページ名 |
| `port` | 接続先ポート（例: `9222`） |
| `Host` | 接続先 IP（例: `127.0.0.1`） |

::: tip 注意
- 必須引数以外をすべて省略すると、`type=page` の先頭タブに繋ぎに行きます
- 内部で「見つかるまでループ」はしません。呼び出し側で別途実装してください
- 接続後は **`CDPContext.reattachWebSocket`** にこのオブジェクトを渡して使います
:::

### `AutoConnectBrowserCDP`

`/json/version` へアクセスし、**ブラウザ単位**の WebSocket 接続まで行います。

```vb
Public Function AutoConnectBrowserCDP(UserName As String, Optional port As Long = 9222, Optional Host As String = "127.0.0.1") As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称 |
| `port` | 接続先ポート（例: `9222`） |
| `Host` | 接続先 IP（例: `127.0.0.1`） |

::: tip 注意
接続後は **`CDPBrowser.reattachWebSocket`** にこのオブジェクトを渡して使います。
:::

::: warning v3.2.0での変更
`ReuseContext` 引数が廃止されました（「Excel テーブルにあるメインタブ情報を流用する」特殊分岐が削除され、接続ロジックが簡略化されています）。
:::

### `ReConnectCDP`（v3.2.0〜）

[`DevToolsActivePort`](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja) ファイルを読んで**今目の前のブラウザ**へ接続するか、記録済みの WinSock ハンドルをそのまま再利用します。

```vb
Public Sub ReConnectCDP(UserName As String, Optional ReuseWinSockHandle As Boolean)
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称 |
| `ReuseWinSockHandle` | `True` で、Excel テーブルに記録済みの WinSock ハンドルをそのまま再利用（無ければエラー）。`False`（既定）は、古いハンドルを破棄したうえで `DevToolsActivePort` ファイルから新規接続 |

::: tip 注意
- 接続後は **`CDPBrowser.reattachWebSocket`** にこのオブジェクトを渡して使います
- `ReuseWinSockHandle:=False` 時、現時点では Edge または Chrome の **安定版** への接続用に限ります
- `ReuseWinSockHandle:=False` の実行直後は、ユーザーが、下記ダイアログに応答するまで、Excelがブロッキングされます
![今目の前のブラウザに接続する際のダイアログ](../public/img/dialog.avif)
:::

::: warning v3.2.0でのリネーム
以前は `AutoConnectDevToolsActivePort(Optional UserName As String = "User Data", Optional ReuseContext As Boolean) As Boolean` という公開メソッドでしたが、v3.2.0で `Private` 化され、代わりに上記の `ReConnectCDP` が公開されました（`ReuseContext`は`ReuseWinSockHandle`に改称）。
:::

## 関連

- [設計思想について](/websocket/design)
- [再接続 (reattach)](/guides/reattach)
- デモ: `Demo_CDP` / `Demo_WebDriverBiDi` / `Demo_WebSocket` の WebSocket 経由セクション
