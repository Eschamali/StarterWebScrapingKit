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

`ShSetting01_StartBrowser.StartCDPMode` / `StartBiDiMode` に `WebSocketMode:=True` を渡すだけで、WebSocket経由でローカルブラウザを起動し、接続済みの `CDPBrowser` / `WebDriverBiDiMode` をそのまま受け取れます。

```vb
Public Function StartCDPMode(Optional StartURL As String, Optional SwitchUser As String, Optional WebSocketMode As Boolean) As CDPBrowser
Public Function StartBiDiMode(Optional StartURL As String, Optional SwitchUser As String, Optional sessionCapabilitiesRequest As Dictionary, Optional WebSocketMode As Boolean) As WebDriverBiDiMode
```

```vb
Sub AutoConnectBrowser()
    '1. WebSocket制御で、ブラウザを起動
    Dim BrowserControl As CDPBrowser
    Set BrowserControl = ShSetting01_StartBrowser.StartCDPMode(WebSocketMode:=True)

    '2. 未接続のタブに接続
    Dim t As CDPContext
    Set t = BrowserControl.getTab(setMain:=True)

    '3. ページ遷移
    t.navigate "https://example.com"

    '4. 終了
    BrowserControl.quit
End Sub
```

`StartCDPModeContext` / `StartBiDiModeContext`（`CDPContext` / `WebDriverBiDiContext` を返す方）には、この引数はありません。WebSocket経由で起動したい場合は `StartCDPMode` / `StartBiDiMode` を使い、`getTab` からタブ操作を始めてください。

### `CDPCoreViaWebSocket` を直接使う場合

内部で何が起きているかを制御したい場合は、`ConnectCDPWithLocalBrowser`（起動＋接続のみ）と `StartCDPMode` / `StartBiDiMode`（接続済みの状態を `CDPBrowser` / `WebDriverBiDiMode` として受け取る）を分けて呼べます。

```vb
Public Sub ConnectCDPWithLocalBrowser( _
    Optional BrowserType As BrowserList = BrowserList.RunChrome, _
    Optional appUrl As String, _
    Optional userProfile As String, _
    Optional addArgs As String _
)
Public Function StartCDPMode() As CDPBrowser
Public Function StartBiDiMode(sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiMode
```

| 引数 | 意味 |
| --- | --- |
| `BrowserType` | `BrowserList` 列挙（`RunChrome` / `RunEdge`） |
| `appUrl` | `--app` に付ける URL |
| `userProfile` | `--user-data-dir` 用のユーザーディレクトリ名 |
| `addArgs` | 追加の起動引数 |

```vb
Dim ws As New CDPCoreViaWebSocket
ws.ConnectCDPWithLocalBrowser BrowserList.RunChrome, "https://example.com"

Dim b As CDPBrowser
Set b = ws.StartCDPMode()

Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
```

内部では、リモートデバッグを禁止するポリシーのチェック・残存セッションの後始末・クラッシュ復元プロンプトの無効化・起動コマンドライン生成・`DevToolsActivePort` の出現待機・接続までを [`CDPHost`](/concepts/architecture) に委託したうえで自動的に行います。`StartCDPMode`/`StartBiDiMode`が返すのは、内部で`reattachWebSocket`済みの`CDPBrowser`/`WebDriverBiDiMode`です。

`Start○○ModeContext`（Pipe版）と違い、返るのは `CDPBrowser` / `WebDriverBiDiMode` です。タブ操作は `getTab` / `newTab` から始めてください。

::: warning v3.1.0での変更
以前このセクションでは `CDPCoreViaWebSocket.RunWebSocketModeBrowserCDP`（起動から`CDPBrowser`取得まで1メソッド）を紹介していましたが、v3.1.0で`ConnectCDPWithLocalBrowser`（起動＋接続）と`StartCDPMode`/`StartBiDiMode`（`CDPBrowser`/`WebDriverBiDiMode`化）に分割されました。日常利用では、上記の設定シート経由（`ShSetting01_StartBrowser.StartCDPMode(WebSocketMode:=True)`）の方が簡単です。
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

`CDPCoreViaWebSocket` には、既存ブラウザへ後付け接続する次の 3 種類のメソッドに加え、前述の起動込みメソッドがあります。

| メソッド | エンドポイント／手段 | 渡す `reattachWebSocket` |
| --- | --- | --- |
| `AutoConnectPageCDP` | `/json/list` → Page | [`CDPContext`](/api/cdp/CDPContext) |
| `AutoConnectBrowserCDP` | `/json/version` → Browser | [`CDPBrowser`](/api/cdp/CDPBrowser) |
| `AutoConnectDevToolsActivePort` | `DevToolsActivePort` ファイル | [`CDPBrowser`](/api/cdp/CDPBrowser) |
| `ConnectCDPWithLocalBrowser` + `StartCDPMode`/`StartBiDiMode` | ローカルブラウザを起動してから接続 | （内部で `reattachWebSocket` 済み。[前述](#ローカルブラウザの起動から行う場合)） |

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
Public Function AutoConnectBrowserCDP( _
    UserName As String, _
    Optional ReuseContext As Boolean, _
    Optional port As Long = 9222, _
    Optional Host As String = "127.0.0.1" _
) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称 |
| `ReuseContext` | `True` で Excel テーブルにあるメインタブ情報を流用 |
| `port` | 接続先ポート（例: `9222`） |
| `Host` | 接続先 IP（例: `127.0.0.1`） |

::: tip 注意
接続後は **`CDPBrowser.reattachWebSocket`** にこのオブジェクトを渡して使います。
:::

### `AutoConnectDevToolsActivePort`

[`DevToolsActivePort`](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja) ファイルを読み、**今目の前のブラウザ**へ接続します。

```vb
Public Function AutoConnectDevToolsActivePort( _
    Optional UserName As String = "User Data", _
    Optional ReuseContext As Boolean _
) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称。省略時は既定の `"User Data"` |
| `ReuseContext` | `True` で Excel テーブルにあるメインタブ情報を流用 |

::: tip 注意
- 接続後は **`CDPBrowser.reattachWebSocket`** にこのオブジェクトを渡して使います
- 現時点では Edge または Chrome の **安定版** への接続用に限ります
- 実行直後は、ユーザーが、下記ダイアログに応答するまで、Excelがブロッキングされます
![今目の前のブラウザに接続する際のダイアログ](../public/img/dialog.avif)
:::

## 関連

- [設計思想について](/websocket/design)
- [再接続 (reattach)](/guides/reattach)
- デモ: `Demo_CDP` / `Demo_WebDriverBiDi` / `Demo_WebSocket` の WebSocket 経由セクション
