# WebSocket モードでできること

ほとんどのブラウザ自動化は Pipe で十分ですが、以下のシチュエーションで自動化する場合は WebSocket モードを使うことになります。

- [Android ブラウザ制御](https://developer.chrome.com/docs/devtools/remote-debugging?hl=ja)
![Androidブラウザが制御されてる様子](../public/viaWebSocket/Android制御.gif)

- [WebView2 制御](https://playwright.dokyumento.jp/docs/webview2)
![WebView2 制御してる様子](../public/viaWebSocket/WebView2制御.avif)

- [今目の前のブラウザを制御](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja)
![今目の前のブラウザを制御してる様子](/viaWebSocket/目の前のブラウザ制御.gif)

- Tailscale 等によるインターネットを介した制御

## 現時点での制限

- ローカル環境でのブラウザ起動メソッドは用意していません。  
手動で `--remote-debugging-port=9222` を付けた状態でローカルブラウザを起動してから、`CDPCoreViaWebSocket.AutoConnectPageCDP` 等を呼び出してください。

## 基本的な接続方法

WebSocket は「後付け」接続のため、Pipe 版の `Start○○ModeContext` とは流れが違います。大まかには次のとおりです。

1. **接続の識別名称を取得／設定** — セル（`ShSetting01_StartBrowser.CurrentUserName`）から取ってもよいし、独自の名前でも OK
2. **目的に合った接続メソッドを呼ぶ** — 下の 3 種類から選択（`CDPCoreViaWebSocket`）
3. **対応する `reattach` に、2. の Class オブジェクトを渡す** — Page 接続なら `CDPContext`、Browser 系なら `CDPBrowser`（または BiDi 側の Mode）
4. **あとはいつも通りの制御**

各 Demo モジュールの **「WebSocket経由版Demo」** セクションを参照してください（`Demo_CDP` / `Demo_WebDriverBiDi` / `Demo_WebSocket`）。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

Dim ws As New CDPCoreViaWebSocket
If Not ws.AutoConnectPageCDP(UserName) Then Exit Sub

Dim t As New CDPContext
If Not t.reattach(UserName, , ws) Then Exit Sub

t.navigate "https://example.com"
ws.DisconnectCDP
```

## 接続の種類

現時点では、`CDPCoreViaWebSocket` に次の 3 種類のメソッドを用意しています。

| メソッド | エンドポイント／手段 | 渡す `reattach` |
| --- | --- | --- |
| `AutoConnectPageCDP` | `/json/list` → Page | [`CDPContext`](/api/cdp/CDPContext) |
| `AutoConnectBrowserCDP` | `/json/version` → Browser | [`CDPBrowser`](/api/cdp/CDPBrowser) |
| `AutoConnectDevToolsActivePort` | `DevToolsActivePort` ファイル | [`CDPBrowser`](/api/cdp/CDPBrowser) |

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
- 接続後は **`CDPContext.reattach`** にこのオブジェクトを渡して使います
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
接続後は **`CDPBrowser.reattach`** にこのオブジェクトを渡して使います。
:::

### `AutoConnectDevToolsActivePort`

[`DevToolsActivePort`](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja) ファイルを読み、**今目の前のブラウザ**へ接続します。

```vb
Public Function AutoConnectDevToolsActivePort( _
    UserName As String, _
    Optional ReuseContext As Boolean _
) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | 利用者識別名称 |
| `ReuseContext` | `True` で Excel テーブルにあるメインタブ情報を流用 |

::: tip 注意
- 接続後は **`CDPBrowser.reattach`** にこのオブジェクトを渡して使います
- 現時点では Edge または Chrome の **安定版** への接続用に限ります
:::

## 関連

- [設計思想について](/websocket/design)
- [再接続 (reattach)](/guides/reattach)
- デモ: `Demo_CDP` / `Demo_WebDriverBiDi` / `Demo_WebSocket` の WebSocket 経由セクション
