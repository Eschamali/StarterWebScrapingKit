# msedge.exe を UserForm に直接埋め込む

> WebView2は、極端な話、Edgeから、URLバーやウィンドウ枠がない状態にしたものみたいなもの。  
> 実はある引数を足してちょこっと簡単なAPIを使えば簡単に実現可能だ。

皆さんのPCに入っている **Microsoft Edge（msedge.exe）そのもの** を、UserFormに埋め込んでしまう力技です。  
見た目は完全に WebView2 コントロール。しかも外部ツール一切不要です。

![EdgeをUserFormに埋め込んだ様子](/img/疑似WebView2.png)

*▲ Google ページが UserForm 内に表示された様子。アドレスバーなし・タブなし = 完全 WebView2 ルック*

## 構成と処理の流れ

Excel と Edge だけで完結する、非常にシンプルでスマートな方式です。

```text
Excel (VBA)  ↔  CDP via Pipe（標準入出力パイプ）  ↔  Edge.exe
```

WebSocket も Node.js も不要。VBA ↔ ブラウザの2者間だけで完全に完結します。

## 実装のポイント

Edge 起動時に、以下の起動引数を付与するのがキモです：

```text
--remote-debugging-pipe --kiosk
```

| 引数 | 効果 |
| --- | --- |
| `--remote-debugging-pipe` | 標準入出力パイプ経由でCDPコマンドを送受信。ポートを開かないため、ネットワーク制限環境でも安全 |
| `--kiosk` | アドレスバー・タブ・ウィンドウ枠が消滅。見た目が完全にWebView2コントロールに |

その後、VBAの **Windows API** を使ってEdgeのウィンドウハンドルを取得し、UserFormの子ウィンドウとして強引に取り込みます。開通したパイプ経由でCDP-JSONコマンドを流し込むことで、スクレイピングや画面操作が可能です。

本キットでは `KioskMode:=fullscreen` などで同等の起動を行います（設定シート／`StartCDPModeContext`）。

## デモコードの動作内容

::: tip
デモでは最初に **このツールのGitHubページ** が表示されます。  
UserForm上部のURLテキストボックスに任意のURL（例：`https://www.google.com`）を入力してEnterを押すと、ページ遷移できることが確認できます。
:::

::: details Demoコードを見る（ExcelのユーザーフォームにEdgeを埋め込む）

```vb
Sub ExcelのユーザーフォームにEdgeを埋め込む()
    '1. CDPでEdgeを起動（Kioskモード = ウィンドウ枠なし）
    Dim 実質WebView2 As CDPContext
    Set 実質WebView2 = ShSetting01_StartBrowser.StartCDPModeContext(KioskMode:=fullscreen)
    実質WebView2.navigate "https://github.com/Eschamali/StarterWebScrapingKit"

    '2. フォームをロード（まだ表示はしない）
    Load EdgeInExcelForm

    '3. 誘拐（ドッキング）処理を実行させる！
    実質WebView2.InheritanceCDPCore.sleep
    If Not (EdgeInExcelForm.AttachEdge(実質WebView2)) Then
        MsgBox "Edgeのハンドル情報の取得に失敗しました", vbCritical
        Exit Sub
    End If

    '4. フォームを表示
    EdgeInExcelForm.show

    '5. ブラウザを正常に閉じる
    実質WebView2.InheritanceCDPBrowser.quit
End Sub
```

同梱デモ: `Demo_CDP.ExcelのユーザーフォームにEdgeを埋め込む`

:::

## デメリット：入力フォーカス問題

::: warning フォーカスが正しく当たらない問題があります
VBAの `UserForm_Activate` イベントは、**UserForm同士の切り替え**でしか動作しません。  
そのため、他のWindowsアプリからUserForm（の中にあるEdge）にフォーカスを戻した際、正しくフォーカスが当たらないことがあります。

現状は「特定領域にマウスを当てて強制的にフォーカス処理を走らせる」という泥臭い工夫で対処しています。
:::

## 次へ

- [PowerShell 経由で真の WebView2](./powershell)
- [総括](./summary)
- [はじめに（UserFormコーナー）](./intro)
