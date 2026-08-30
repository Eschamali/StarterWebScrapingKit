---
description: Excel の UserForm に埋め込んだ WebView2 を、Pipe / WebSocket と同じ CDPContext / CDPElement の感覚で制御する方法を紹介します。
---

# WebView2モードでできること

Pipe・WebSocket が「外部のブラウザプロセス」を相手にするのに対し、WebView2 モードは **Excel自身のUserFormに埋め込んだブラウザ**を制御します。デバッグポートも名前付きパイプも使いません。

- Excel の UserForm に本物の WebView2 を埋め込み、リッチな画面（React / Vue / SPA など）を表示しつつ、同じタブを CDP で操作したい場合
- 社内ツールとして、1枚の xlsm だけで「ブラウザ埋め込みUI」を配布したい場合

詳しい経緯・実装の考え方は [設計思想について](/webview2/design) を参照してください。

## 基本的な接続方法

いちばん簡単なのは、同梱の `WebView2Form` を使う方法です。

```vb
Sub ExcelのユーザーフォームにWebView2を埋め込む()
    With WebView2Form
        If Not .StartCDPModeWebView2 Then Debug.Print "WebView2の初期化に失敗しました。": Exit Sub

        .ThisCDPContext.navigate "https://github.com/Eschamali/StarterWebScrapingKit"
        .show
    End With
End Sub
```

同梱デモ: `Demo_CDP.ExcelのユーザーフォームにWebView2を埋め込む`

内部では、`WebView2Form.StartCDPModeWebView2` が `CDPCoreViaWebView2.ConnectCDP` を呼んでWebView2の`Environment`/`Controller`/`ICoreWebView2`を生成し、`CDPBrowser.reattachWebView2` / `CDPContext.reattachWebView2` を通じて、Pipe版・WebSocket版と**まったく同じCDPスタック**に接続します。埋め込んでしまえば、`getElementByQuery` や `jsEval` など、これまでのガイドで説明してきた操作がそのまま使えます。

## 自前のUserFormに組み込む場合

`CDPCoreViaWebView2` を直接使えば、自作のUserFormにも組み込めます。

```vb
Public Function ConnectCDP(UserName As String, Optional AttachHwnd As LongPtr) As Boolean
```

| 引数 | 意味 |
| --- | --- |
| `UserName` | WebView2 のユーザーデータフォルダ名 |
| `AttachHwnd` | WebView2 を貼り付けるウィンドウハンドル。省略時は Excel 自身のハンドル（`Application.Hwnd`）を使用 |

```vb
Dim wv2 As New CDPCoreViaWebView2
If Not wv2.ConnectCDP("MyUser", Me.EdgeFrame.hWnd) Then Exit Sub

Dim b As New CDPBrowser
b.reattachWebView2 "MyUser", wv2

Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
```

::: tip 注意
- 既存の接続がある場合は、`ConnectCDP` の再呼び出しで切断・再接続されます
- `CDPBrowser.newTab`（`Target.createTarget`）自体はWebView2モードでも使えますが、WebView2は1インスタンス=1ページのため、新規タブはUserForm内には埋め込まれず**独立した新規ウィンドウ**として開きます。タブ（ウィンドウ）をまたいだCDPコマンドのやり取りは`CallDevToolsProtocolMethodForSession`が担うため、複数の`CDPContext`を並行操作すること自体は可能です（詳細は[設計思想について](/webview2/design)）
:::

## 表示・イベント購読

| メンバー | 役割 |
| --- | --- |
| `Resize(Width, Height, Optional Top, Optional Left)` | WebView2の表示サイズ・位置を変更 |
| `Visible`（Let） | 表示/非表示の切り替え |
| `DevToolsEnabled`（Let） | 右クリックの「検証」等、開発者ツールの有効/無効 |
| `ContextMenuEnabled`（Let） | 右クリックメニューの有効/無効 |
| `SubscribeCdpEvent(EventName) As Boolean` | 指定したCDPイベント名を個別に購読開始 |
| `UnsubscribeCdpEvent(EventName) As Boolean` | 指定したCDPイベント名の購読を解除 |
| `UnsubscribeAllCdpEvents() As Long` | 購読中の全イベントを一括解除（解除件数を返す） |
| `SubscribeCdpEventCount`（Get） | 購読中のイベント数 |
| `isAvailability`（Get） | WebView2（`ICoreWebView2`）が生きているか |

```vb
' Page.loadEventFired を購読してから遷移する例
wv2.SubscribeCdpEvent "Page.loadEventFired"
t.navigate "https://example.com"
' ... TakeEvents ループ等で受信 ...
wv2.UnsubscribeAllCdpEvents
```

::: warning WebSocket/Pipeとの違い
Pipe / WebSocket は「ドメインを`enable`すれば、そのドメインの全イベントが自動で流れてくる」モデルですが、WebView2は`GetDevToolsProtocolEventReceiver`の仕様上、**イベント名ごとの個別購読**が必要です。一括購読の概念はWebView2側に無いため未対応です（一括解除のみ`UnsubscribeAllCdpEvents`として提供）。
:::

## 再接続 (reattach)

Pipe / WebSocket と同じく、`reattachWebView2` で既存のWebView2接続情報に再接続できます。

```vb
Public Function reattachWebView2(userProfile As String, WebView2Mode As CDPCoreViaWebView2, Optional reuseSession As Boolean) As Boolean
```

詳細は [再接続 (reattach)](/guides/reattach) を参照してください。

## 関連

- [設計思想について](/webview2/design) — 機械語サンク・vtable、移植元へのクレジット
- [Excel単独で「真のWebView2」を完全制御する](/userform/vba-only) — UserForm埋め込みの詳しい解説
- [再接続 (reattach)](/guides/reattach)
- デモ: `Demo_CDP.ExcelのユーザーフォームにWebView2を埋め込む`
