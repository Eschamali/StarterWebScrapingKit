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

SessionId を保持したい場合は Part1 側で `KeepSession = True`（デモコメント参照）。

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
* 既存ブラウザ（デバッグポート）への接続は `CDPCoreViaWebSocket` を `reattach` に渡すパターン（`Demo_CDP.AutoConnect*`）
:::

## 関連デモ

- `Demo_CDP.demoReattachmentPart*`
- `Demo_WebDriverBiDi.demoReattachmentPart*`
