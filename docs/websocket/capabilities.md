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

**ブラウザを初めて立ち上げる場合**は、ブラウザ起動設定シートの **`UseWebSocket` セルを `TRUE`** にするだけで、Pipe の代わりに WebSocket 経由でローカルブラウザを起動・接続するようになります（既定は `FALSE` = Pipe）。

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
    BrowserControl.CloseBrowser
End Sub
```

*▲ 書き方は、`Pipe`版と変わりません🤠*

## 既に起動中のデバッグ中ブラウザを制御する場合

すでにデバッグポート付きで起動しているブラウザには、`CDPCoreViaWebSocket` の接続メソッドで WebSocket 接続を確立し、その接続を [`ReattachOptions`](/api/cdp/ReattachOptions) の `Set_WebSocketMode` に渡して、`reattach` で後付け接続できます。

| 接続メソッド（`CDPCoreViaWebSocket`） | 接続先 | 渡す `reattach` |
| --- | --- | --- |
| `AutoConnectPageCDP` | Page | [`CDPContext.reattach`](/api/cdp/CDPContext) |
| `AutoConnectBrowserCDP` | Browser | [`CDPBrowser.reattach`](/api/cdp/CDPBrowser) / `WebDriverBiDiMode.reattach` |
| `AutoConnectDevToolsActivePort` | 今目の前のブラウザ（`DevToolsActivePort`） | [`CDPBrowser.reattach`](/api/cdp/CDPBrowser) / `WebDriverBiDiMode.reattach` |
| `ReConnectCDP` | 記録済みの WinSock ハンドルを再利用 | [`CDPBrowser.reattach`](/api/cdp/CDPBrowser) / `WebDriverBiDiMode.reattach` |

各メソッドの詳細は「接続の種類」を参照してください。

### CDP で制御する例

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

'1. 起動中のブラウザへ接続（Page 単位なら AutoConnectPageCDP、今目の前のブラウザなら AutoConnectDevToolsActivePort）
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket        ' ※`As New` で宣言しないこと
If Not ws.AutoConnectBrowserCDP(UserName) Then Exit Sub

'2. 接続済みの WebSocket を、制御経路として指定
ReattachOptions.Set_WebSocketMode(ws) = UserName

'3. CDPBrowser として受け取る
Dim b As New CDPBrowser
b.reattach ReattachOptions

'4. あとはいつも通りの制御
Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
```

### BiDi で制御する例

```vb
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket
If Not ws.AutoConnectBrowserCDP(UserName) Then Exit Sub

ReattachOptions.Set_WebSocketMode(ws) = UserName

Dim bidi As New WebDriverBiDiMode
bidi.reattach ReattachOptions
```

## 基本的な接続方法（既存ブラウザへの後付け接続）

WebSocket は「後付け」接続のため、Pipe 版の `Start○○ModeContext` とは流れが違います。大まかには次のとおりです。

1. **接続の識別名称を取得／設定** — セル（`ShSetting01_StartBrowser.CurrentUserName`）から取ってもよいし、独自の名前でも OK
2. **目的に合った接続メソッドを呼ぶ** — 下の接続メソッドから選択（`CDPCoreViaWebSocket`）
3. **接続した Class オブジェクトを `ReattachOptions.Set_WebSocketMode(ws) = UserName` で指定し、`reattach` に渡す** — Page 接続なら `CDPContext`、Browser 系なら `CDPBrowser`（または BiDi 側の Mode の `reattach`）
4. **あとはいつも通りの制御**

各 Demo モジュールの **「WebSocket経由版Demo」** セクションを参照してください（`Demo_CDP` / `Demo_WebDriverBiDi`）。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket
If Not ws.AutoConnectPageCDP(UserName) Then Exit Sub

ReattachOptions.Set_WebSocketMode(ws) = UserName

Dim t As New CDPContext
If Not t.reattach(ReattachOptions) Then Exit Sub

t.navigate "https://example.com"
t.ThisCDPBrowser.DisconnectCDP
```

## 接続の種類

`CDPCoreViaWebSocket` には、既存ブラウザへ後付け接続する次の公開メソッドがあります。

| メソッド | エンドポイント／手段 | 接続先 |
| --- | --- | --- |
| `AutoConnectPageCDP` | `/json/list` → Page | [`CDPContext`](/api/cdp/CDPContext) |
| `AutoConnectBrowserCDP` | `/json/version` → Browser | [`CDPBrowser`](/api/cdp/CDPBrowser) / `WebDriverBiDiMode` |
| `AutoConnectDevToolsActivePort` | `DevToolsActivePort` ファイル | [`CDPBrowser`](/api/cdp/CDPBrowser) / `WebDriverBiDiMode` |
| `ReConnectCDP` | 記録済みハンドルの再利用 | [`CDPBrowser`](/api/cdp/CDPBrowser) / `WebDriverBiDiMode` |

ローカルブラウザの起動から接続までを行う `OpenBrowserAndConnectCDP(caps)` もありますが、通常は設定シートの「WebSocketモード」と `StartCDPMode` などの起動ヘルパー経由で足ります（前述）。

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
- 接続後は **`ReattachOptions.Set_WebSocketMode(ws) = UserName`** で指定し、**`CDPContext.reattach`** に渡して使います
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
接続後は **`ReattachOptions.Set_WebSocketMode(ws) = UserName`** で指定し、**`CDPBrowser.reattach`** に渡して使います。
:::

### `AutoConnectDevToolsActivePort`

[`DevToolsActivePort`](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja) ファイルを読んで、**今目の前のブラウザ**へ接続します。

```vb
Public Function AutoConnectDevToolsActivePort(caps As BrowserCapabilities) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `caps` | `DevToolsActivePort` ファイルを探すための情報をまとめた [`BrowserCapabilities`](/api/cdp/BrowserCapabilities)。対象のユーザーデータフォルダ名は `Set_UserDataDirName` で指定します |

```vb
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket

BrowserCapabilities.Set_UserDataDirName = "User Data"
ws.AutoConnectDevToolsActivePort BrowserCapabilities

ReattachOptions.Set_WebSocketMode(ws) = "User Data"

Dim b As New CDPBrowser
b.reattach ReattachOptions
```

::: tip 注意
- 接続後は **`ReattachOptions.Set_WebSocketMode(ws) = UserName`** で指定し、**`CDPBrowser.reattach`** に渡して使います
- 既にデバッグブラウザが起動済みの前提で動くため、`DevToolsActivePort` ファイルの存在チェックは行いません
- 実行直後は、ユーザーが、下記ダイアログに応答するまで、Excelがブロッキングされます
![今目の前のブラウザに接続する際のダイアログ](../public/img/dialog.avif)
:::

::: tip 応用
**任意パスの `DevToolsActivePort` を読む**: `Set_UserDataDirName` に絶対パス形式を渡すと、そのパスの `DevToolsActivePort` ファイルを読み込んで接続できます
:::

### `ReConnectCDP`

記録済みの WinSock ハンドルをそのまま適用して、WebSocket の再接続儀式を飛ばします。

```vb
Public Sub ReConnectCDP(UserName As String)
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | どのユーザーの WinSock ハンドルを使うか。Excel テーブルに記録が無い場合は VBA エラー |

::: tip 接続確認ダイアログをスキップ
初回接続（`AutoConnectDevToolsActivePort`）を済ませたあと、WebSocket を切断せずにプロシージャを終え、次回以降に同じユーザー名で `ReConnectCDP` を使うと、接続確認ダイアログを出さずに再接続できます。このとき、`ReattachOptions.Set_Destruction = False`（古いデータを破棄しない）にすることを推奨します。

```vb
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket
ws.ReConnectCDP "User Data"

With ReattachOptions
    .Set_WebSocketMode(ws) = "User Data"
    .Set_Destruction = False
End With

Dim b As New CDPBrowser
b.reattach ReattachOptions
```

![ReConnectCDP で、接続確認ダイアログを出さずに再接続している様子](/viaWebSocket/wow.gif)

*▲ `Demo_CDP.AutoConnectDevToolsActivePort` を使い、WinSock ハンドルを再利用して再接続している様子。接続確認ダイアログが出ることなく、つながります*
:::

## 関連

- [設計思想について](/websocket/design)
- [再接続 (reattach)](/guides/reattach)
- [`ReattachOptions`](/api/cdp/ReattachOptions) / [`BrowserCapabilities`](/api/cdp/BrowserCapabilities)
- デモ: `Demo_CDP` / `Demo_WebDriverBiDi` の WebSocket 経由セクション
