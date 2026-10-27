---
description: 起動済みブラウザへの再接続（reattach）時のオプションを1つにまとめる設定保持専用クラス ReattachOptions。制御経路の指定と、再接続時の挙動を設定します。
---

# ReattachOptions

起動済みのブラウザへ**再接続（reattach）するときのオプション設定を、1つにまとめたクラス**です。`CDPBrowser.reattach` / `CDPContext.reattach` / `WebDriverBiDiMode.reattach` / `WebDriverBiDiContext.reattach` の引数に渡します。

起動時のオプションをまとめた [`BrowserCapabilities`](./BrowserCapabilities) の再接続版です。再接続では特に、**どの制御経路（Pipe / Port / WebView2）でつなぎ直すか**を必ず指定する必要があります。

- `BrowserCapabilities` と同じく、起動に使ったあとは破棄される **使い切り（ワンショット）** です
- `ReattachOptions` という名前そのものが、そのまま使える**既定インスタンス**です（`New` 不要）

## 使い方

制御経路を指定して、`reattach` に渡します。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

' 1. 制御経路を指定する（Pipe の例）
ReattachOptions.Set_PipeMode = UserName

' 2. reattach に渡す
Dim b As New CDPBrowser
b.reattach ReattachOptions

' 3. あとは通常どおり
Dim t As CDPContext
Set t = b.getTab(setMain:=True)
t.navigate "https://example.com"
```

### 渡す先のメソッド

| クラス | メソッド | 戻り値 |
| --- | --- | --- |
| [`CDPBrowser`](./CDPBrowser) | `reattach(reattachInfo As ReattachOptions)` | なし（失敗時は VBA エラー） |
| [`CDPContext`](./CDPContext) | `reattach(reattachInfo As ReattachOptions, Optional reuseSession As Boolean) As Boolean` | 既存タブへの接続成功可否 |
| [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) | `reattach(reattachInfo As ReattachOptions)` | なし（失敗時は VBA エラー） |
| [`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) | `reattach(reattachInfo As ReattachOptions) As Boolean` | 既存 context への接続成功可否 |

詳しい流れは [再接続ガイド](/guides/reattach) を参照してください。

## 制御経路の設定

次の3つのうち、**いずれか1つを必ず指定**します。どれも指定せずに `reattach` へ渡すと、VBA エラーになります。複数を指定した場合は、最後に指定したものだけが有効です（内部の通信コアが作り直されるため）。

いずれも、識別名称（`UserName`）に紐づく接続情報を Excel テーブルから読み取ります。該当する行が無い場合は、その場で VBA エラーになります。

| プロパティ | 経路 | 使い方 |
| --- | --- | --- |
| `Set_PipeMode` | Pipe（`--remote-debugging-pipe`） | `ReattachOptions.Set_PipeMode = UserName` |
| `Set_WebSocketMode` | Port（`--remote-debugging-port`） | `ReattachOptions.Set_WebSocketMode(ws) = UserName`（`ws` は接続済みの `CDPCoreViaWebSocket`） |
| `Set_WebView2Mode` | WebView2 | `ReattachOptions.Set_WebView2Mode(wv) = UserName`（`wv` は接続済みの `CDPCoreViaWebView2`） |

```vb
' Port の例：既存ブラウザへ WebSocket 接続してから、その接続を渡す
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket        ' ※`As New` で宣言しないこと
Debug.Print ws.AutoConnectPageCDP(UserName)

ReattachOptions.Set_WebSocketMode(ws) = UserName
```

::: warning `Dim ws As New CDPCoreViaWebSocket` は使わない
`As New` で宣言すると、`Nothing` 判定のたびに自動で `New` されてしまうため、接続オブジェクトの受け渡しで意図しない挙動になります。`Dim ws As CDPCoreViaWebSocket` と `Set ws = New CDPCoreViaWebSocket` に分けて書いてください。
:::

## その他の設定

| プロパティ | 型 | 既定値 | 意味 |
| --- | --- | --- | --- |
| `Set_Destruction` | `Boolean` | `True` | 再接続後に、通信バッファーに溜まっている古いデータを破棄するか。`False` にすると、破棄せずにイベント情報を受け取ります（WinSock ハンドルを使いまわす場合など） |
| `Set_EnableUserAgentOverride` | `Boolean` | シートのデバッグ設定 | ブラウザ本来の `User-Agent` の末尾に、このツールのサインを追記するか |

### WebDriver BiDi 用

BiDi の再接続で、BiDi-CDP Mapper が再始動された場合に使われます。

| プロパティ | 型 | 既定値 | 意味 |
| --- | --- | --- | --- |
| `Set_BiDiSessionCapabilitiesRequest` | `Dictionary`（`Property Set`） | 未設定 | 再始動後の `session.new` に渡す `capabilities` |
| `Set_BiDiViewTab_BiDiCDPMapper` | `Boolean` | シートのデバッグ設定 | Mapper のコンソールタブを表示するか |
| `Set_BiDiViewDebugLogForBiDiCDPMapper` | `Boolean` | シートのデバッグ設定 | Mapper 内部のデバッグログを記録するか |

## 使い切り（ワンショット）の仕様

`reattach` を呼ぶと、このオブジェクトは自動で破棄されます（既定インスタンスの `ReattachOptions` も含みます）。次に `ReattachOptions` を使うと新しいインスタンスとして再生成され、制御経路は**未設定の状態**から始まります。`reattach` のたびに、制御経路を指定し直してください。

```vb
ReattachOptions.Set_PipeMode = UserName
b.reattach ReattachOptions            ' ここで破棄される

' もう一度 reattach するなら、経路の指定からやり直す
ReattachOptions.Set_PipeMode = UserName
t.reattach ReattachOptions
```

## 関連

- [再接続 (reattach)](/guides/reattach)
- [`BrowserCapabilities`](./BrowserCapabilities) — 起動用のオプションオブジェクト
- [WebSocket モード](/websocket/capabilities) / [WebView2 モード](/webview2/capabilities)
