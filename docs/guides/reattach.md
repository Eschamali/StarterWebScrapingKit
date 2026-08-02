---
description: プロシージャをまたいで同じデバッグブラウザへ戻る reattach。パイプ／セッションの復元と KeepSession の使い方を解説します。
---

# 再接続 (reattach)

認証の手作業など、**プロシージャをまたいで**同じデバッグブラウザへ戻りたいときに使います。

## 流れ

1. Part1: いつもどおり起動・ナビして処理を中断（ブラウザは開いたまま）
2. Part2: `reattach` でパイプ情報／コンテキストを復元し、続きを実行

## CDP — ブラウザ単位

```vb
' --- Part1 ---
Dim c As CDPContext
Set c = ShSetting01_StartBrowser.StartCDPModeContext
c.navigate "https://google.com"

' --- Part2（別プロシージャ）---
Dim b As New CDPBrowser
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

If Not b.reattach(UserName) Then
    MsgBox "接続できません"
    Exit Sub
End If

Dim r As CDPContext
Set r = b.getTab(setMain:=True)   ' 必須: setMain:=True
r.navigate "https://example.com"
```

## CDP — タブ（Context）単位

```vb
Dim c As New CDPContext
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

' 第2引数: Excel に記録した SessionId を再利用するか
If Not c.reattach(UserName, False) Then
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
2. 再開したいプロシージャで、タブ Class の `reattach` を `CDPContext.reattach(UserName, True)` として呼ぶ
3. あとは今まで通り

```vb
' --- Part1 ---
Dim c As CDPContext
Set c = ShSetting01_StartBrowser.StartCDPModeContext
' ... JavaScript 実行などで ObjectID を得る ...
c.KeepSession = True

' --- Part2（別プロシージャ）---
Dim t As New CDPContext
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

If Not t.reattach(UserName, True) Then Exit Sub
' 同じ SessionID 上で ObjectID を引き続き利用可能
```

### WebSocket 版

1. 処理の最後に `CDPContext.KeepSession = True` にしつつ、`CDPCoreViaWebSocket.DisconnectCDP` などの**切断処理をしない**
2. 利用者識別名称（`UserName`）で `CDPCoreViaWebSocket.deserialize(UserName)` し、WinSock ハンドルを復元
3. 再開したいプロシージャで、タブ Class の `reattach` を `CDPContext.reattach(UserName, True, WebSocketCDP)` として呼ぶ
4. あとは今まで通り

```vb
' --- Part1 ---
' AutoConnectPageCDP → reattach 後の処理 ...
t.KeepSession = True
' ※ ここでは DisconnectCDP しない

' --- Part2（別プロシージャ）---
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName

Dim WebSocketCDP As New CDPCoreViaWebSocket
If WebSocketCDP.deserialize(UserName) <> 0 Then Exit Sub

Dim t As New CDPContext
If Not t.reattach(UserName, True, WebSocketCDP) Then Exit Sub
```

詳細は [WebSocket モードでできること](/websocket/capabilities) も併せて参照してください。

## BiDi — Mode / Context

```vb
' Part1
Dim First As WebDriverBiDiContext
Set First = ShSetting01_StartBrowser.StartBiDiModeContext
First.navigate "https://www.google.com/"

' Part2 — Mode
Dim mode As New WebDriverBiDiMode
Dim UserName As String
UserName = ShSetting01_StartBrowser.CurrentUserName
If Not mode.reattach(UserName) Then Exit Sub

Dim tab As WebDriverBiDiContext
Set tab = mode.getTab(setMain:=True)
If tab Is Nothing Then
    MsgBox "有効なタブがありません。タブを追加して再試行"
    Exit Sub
End If
tab.navigate "https://example.com"

' Part2 — 最後に操作した Context 直接
Dim ctx As New WebDriverBiDiContext
If Not ctx.reattach(UserName) Then Exit Sub
ctx.navigate "https://w3c.github.io/webdriver-bidi/"
```

::: warning 注意
* パイプハンドルや Target / BiDi context が死んでいると失敗します。その場合は Part1 からやり直し
* BiDi の mapper タブが消えても、`WebDriverBiDiMode.reattach` で再始動できる場合があります
* 既存ブラウザ（デバッグポート）への接続は `CDPCoreViaWebSocket` を `reattach` に渡すパターン（`Demo_CDP.AutoConnect*`）。詳しくは [WebSocket モード](/websocket/capabilities)
:::

## 関連デモ

- `Demo_CDP.demoReattachmentPart*`
- `Demo_WebDriverBiDi.demoReattachmentPart*`
