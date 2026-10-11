---
description: プロシージャをまたいで同じデバッグブラウザへ戻る reattach。ReattachOptions で制御経路を指定し、パイプ／WebSocket／WebView2 の接続と SessionID を引き継ぐ方法を解説します。
---

# 再接続 (reattach)

認証の手作業など、**プロシージャをまたいで**同じデバッグブラウザへ戻りたいときに使います。

## 流れ

1. Part1: いつもどおり起動・ナビして処理を中断（ブラウザは開いたまま）
2. Part2: [`ReattachOptions`](/api/cdp/ReattachOptions) で制御経路を指定し、`reattach` でパイプ情報／コンテキストを復元して続きを実行

`reattach` に渡す `ReattachOptions` は、次の3つのうち**いずれか1つの制御経路を必ず指定**します（指定しないと VBA エラーになります）。

| 制御経路 | 指定の書き方 |
| --- | --- |
| Pipe | `ReattachOptions.Set_PipeMode = UserName` |
| Port（WebSocket） | `ReattachOptions.Set_WebSocketMode(ws) = UserName` |
| WebView2 | `ReattachOptions.Set_WebView2Mode(wv) = UserName` |

::: tip 使い切りです
`ReattachOptions` は `reattach` を呼ぶと破棄されます。`reattach` のたびに、制御経路を指定し直してください。
:::

## CDP — ブラウザ単位

`CDPBrowser.reattach` に渡します。ここでは基本の Pipe 版を示します。

```vb
' --- Part1 ---
Dim c As CDPContext
Set c = ShSetting01_StartBrowser.StartCDPModeContext
c.navigate "https://google.com"

' --- Part2（別プロシージャ）---
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

ReattachOptions.Set_PipeMode = UserName      ' 制御経路を指定

Dim b As New CDPBrowser
b.reattach ReattachOptions                   ' パイプが死んでいる場合はここで VBA エラー停止

Dim r As CDPContext
Set r = b.getTab(setMain:=True)              ' 必須: setMain:=True
r.navigate "https://example.com"
```

::: tip 注意
`CDPBrowser.reattach` は戻り値なしの `Sub` です。接続情報が生きていない場合は `Boolean` を返さず VBA エラーで停止します。エラーで止めたくない場合は `On Error` を使ってください。
:::

## CDP — タブ（Context）単位

`CDPContext.reattach` は `Function ... As Boolean` で、失敗時は例外を投げず `False` を返します。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

ReattachOptions.Set_PipeMode = UserName

' 第2引数: Excel に記録した SessionId を再利用するか
Dim c As New CDPContext
If Not c.reattach(ReattachOptions, False) Then
    MsgBox "TargetID が無効です"
    Exit Sub
End If
c.navigate "https://example.com"
```

SessionID を保持して引き継ぐ場合は、次節を参照してください。

## SessionID の引き継ぎについて

基本的には、SessionID は再更新して繋ぎ直すのが CDP のルールです。ただし次のようなシチュエーションでは、SessionID を保持しておくとよいでしょう。

- **JavaScript 実行結果の ObjectID を次のプロシージャでも使いまわしたい場合**  
  → SessionID を更新するとこれらが失われるため

::: warning
この SessionID 引き継ぎは、CDP の [`CDPContext.reattach`](/api/cdp/CDPContext) からのみ対応します。
:::

### Pipe 版

1. 処理の最後に `CDPContext.KeepSession = True`
2. 再開したいプロシージャで、`ReattachOptions.Set_PipeMode` を指定したうえで、`CDPContext.reattach(ReattachOptions, True)` として呼ぶ
3. あとは今まで通り

```vb
' --- Part1 ---
Dim c As CDPContext
Set c = ShSetting01_StartBrowser.StartCDPModeContext
' ... JavaScript 実行などで ObjectID を得る ...
c.KeepSession = True

' --- Part2（別プロシージャ）---
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName
ReattachOptions.Set_PipeMode = UserName

Dim t As New CDPContext
If Not t.reattach(ReattachOptions, True) Then Exit Sub
' 同じ SessionID 上で ObjectID を引き続き利用可能
```

### Port（WebSocket）版

Port は「後付け」で繋ぐ経路のため、既存ブラウザへの接続処理（`AutoConnectPageCDP` など）を行ってから、その接続を `ReattachOptions` に渡します。

```vb
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

' 1. 既存ブラウザへ WebSocket 接続
Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket          ' ※`As New` で宣言しないこと
If Not ws.AutoConnectPageCDP(UserName) Then Exit Sub

' 2. 接続した WebSocket を、制御経路として指定
ReattachOptions.Set_WebSocketMode(ws) = UserName

' 3. reattach
Dim t As New CDPContext
If Not t.reattach(ReattachOptions) Then Exit Sub
```

詳しくは [WebSocket モードでできること](/websocket/capabilities) を参照してください。

#### WinSock ハンドルの引き継ぎ（接続処理を飛ばす）

1. 処理の最後に `CDPContext.KeepSession = True` にしつつ、`DisconnectCDP` などの**切断処理をしない**
2. 再開したいプロシージャで、`CDPCoreViaWebSocket.ReConnectCDP(UserName)` により、Excel に記録済みの WinSock ハンドルを復元する
3. `ReattachOptions.Set_WebSocketMode(ws) = UserName` に加え、`ReattachOptions.Set_Destruction = False`（古いデータを破棄しない）にして、`CDPContext.reattach(ReattachOptions, True)` を呼ぶ

```vb
' --- Part1 ---
' AutoConnectPageCDP → reattach 後の処理 ...
t.KeepSession = True
' ※ ここでは DisconnectCDP しない

' --- Part2（別プロシージャ）---
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

Dim ws As CDPCoreViaWebSocket
Set ws = New CDPCoreViaWebSocket
ws.ReConnectCDP UserName                  ' 記録済みの WinSock ハンドルを復元（無ければ VBA エラー）

With ReattachOptions
    .Set_WebSocketMode(ws) = UserName
    .Set_Destruction = False              ' WinSock を使いまわすときは False
End With

Dim t As New CDPContext
If Not t.reattach(ReattachOptions, True) Then Exit Sub
```

::: tip 今目の前のブラウザへ接続する場合
`edge://inspect/#remote-debugging` で許可した、今目の前のブラウザへ繋ぐ場合は、`ReConnectCDP` の代わりに `CDPCoreViaWebSocket.AutoConnectDevToolsActivePort(BrowserCapabilities)` で接続します（対象のユーザーデータフォルダ名は `BrowserCapabilities.Set_UserDataDirName` で指定）。詳細は [WebSocket モードでできること](/websocket/capabilities) を参照してください。
:::

### WebView2 版

考え方は同じです。接続済みの `CDPCoreViaWebView2` を `ReattachOptions.Set_WebView2Mode(wv) = UserName` で指定し、`CDPContext.reattach(ReattachOptions, True)` を呼びます。詳細は [WebView2モードでできること](/webview2/capabilities) / [UserForm への埋め込み（Lv.99: Excel単体）](/userform/vba-only) を参照してください。

## BiDi — Mode / Context

```vb
' Part1
Dim First As WebDriverBiDiContext
Set First = ShSetting01_StartBrowser.StartBiDiModeContext
First.navigate "https://www.google.com/"

' Part2 — Mode
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName
ReattachOptions.Set_PipeMode = UserName

Dim mode As New WebDriverBiDiMode
mode.reattach ReattachOptions                ' 失敗時は VBA エラーで停止

Dim tab As WebDriverBiDiContext
Set tab = mode.getTab(setMain:=True)
If tab Is Nothing Then
    MsgBox "有効なタブがありません。タブを追加して再試行"
    Exit Sub
End If
tab.navigate "https://example.com"

' Part2 — 最後に操作した Context 直接
ReattachOptions.Set_PipeMode = UserName

Dim ctx As New WebDriverBiDiContext
If Not ctx.reattach(ReattachOptions) Then Exit Sub
ctx.navigate "https://w3c.github.io/webdriver-bidi/"
```

::: warning 注意
* パイプハンドルや Target / BiDi context が死んでいると失敗します。その場合は Part1 からやり直し
* `WebDriverBiDiMode.reattach` は戻り値なしの `Sub`（失敗時は VBA エラー）、`WebDriverBiDiContext.reattach` は `Boolean` を返します
* BiDi の mapper タブが消えても、`WebDriverBiDiMode.reattach` で再始動できる場合があります。このとき、`ReattachOptions.Set_BiDiSessionCapabilitiesRequest` などの BiDi 用の設定が使われます
* 既存ブラウザ（デバッグポート）への接続は、CDP でも BiDi でも、接続済みの `CDPCoreViaWebSocket` を `ReattachOptions.Set_WebSocketMode(ws)` で指定するパターンです（`Demo_CDP.AutoConnect*` / `Demo_WebDriverBiDi.AutoConnect*`）。詳しくは [WebSocket モード](/websocket/capabilities)
:::

「BiDiのmapperタブは消えているがブラウザ自体は生きている」ケースで再接続しても、Excel テーブルに記録済みの `BiDi-context`（メインタブ）情報は保持されます。

## 関連デモ

- `Demo_CDP.demoReattachmentPart*`
- `Demo_WebDriverBiDi.demoReattachmentPart*`
