---
description: Excel VBA でブラウザ自動化を始める最短手順。マクロブックの信頼設定から CDP / BiDi の Hello World、起動モードの選び方までを解説します。
---

# はじめに

最短でブラウザを動かすまでの手順です。

## 1. マクロブックを信頼する

インターネットからダウンロードしたファイルには Mark of the Web (MOTW) が付きます。次を行ってください。

1. Excel をすべて閉じる
2. ファイルを右クリック → **プロパティ**
![右クリックメニュー](/img/GettingStarted/FirstStep4.png)

3. **許可する** にチェック → OK
![プロパティウィンドウ](/img/GettingStarted/FirstStep5.png)

4. 再度開き、「編集を有効にする」／マクロを許可
![マクロ警告バー](/img/GettingStarted/FirstStep6.png)

詳細な説明はリポジトリの [`README-jp.md`](https://github.com/Eschamali/StarterWebScrapingKit/blob/main/README-jp.md) を参照してください。

## 2. ブラウザ起動設定シート

ワークシート **ブラウザ起動設定ver2.X.X** で次を確認します。

- ユーザーデータフォルダ名（`user-data-dir` 用）
- 追加の起動引数（基本は、J13セル以降に記述）
- 起動時のブラウザ表示モード（[ShowWindowのnCmdShow](https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-showwindow)に準拠）
- 特定のChromiumブラウザで自動化する場合はそのフルパス
![基本設定画面](/img/GettingStarted/SettingGUI1.png)


こだわりがなければシートの初期値のままで問題ありません。

## 3. Hello World

標準モジュールのデモ、または新規モジュールに次を書いて実行します。

::: code-group

```vb [CDP]
Sub CDPによる冒険の始まり()
    Dim HelloWorld As CDPContext
    Set HelloWorld = ShSetting01_StartBrowser.StartCDPModeContext

    HelloWorld.navigate "https://kemono-friends.jp/"
    HelloWorld.notify "あなたは、けものがお好きですか？"
    HelloWorld.InheritanceCDPBrowser.sleep 3

    HelloWorld.InheritanceCDPBrowser.quit
End Sub
```

```vb [BiDi]
Sub BiDiによる冒険の始まり()
    Dim HelloWorld As WebDriverBiDiContext
    Set HelloWorld = ShSetting01_StartBrowser.StartBiDiModeContext

    HelloWorld.navigate "https://example.com"

    HelloWorld.InheritanceWebDriverBiDiMode.quit
End Sub
```

:::

同梱デモ:

- `Demo_CDP.CDPによる冒険の始まり`
- `Demo_WebDriverBiDi.BiDiによる冒険の始まり`

### 起動ヘルパーメソッド

シートオブジェクト：ShSetting01_StartBrowser を利用して少ない引数で直ぐに開始できるヘルパーメソッドです。  
下記4種類を用意しております。

| 関数 | 戻り値 | 用途 |
| --- | --- | --- |
| `StartCDPModeContext` | `CDPContext` | タブ操作からすぐ始めたい（推奨） |
| `StartCDPMode` | `CDPBrowser` | タブを自分で `newTab` / `getTab` したい |
| `StartBiDiModeContext` | `WebDriverBiDiContext` | BiDi でタブ操作から始める（推奨） |
| `StartBiDiMode` | `WebDriverBiDiMode` | BiDi でセッション／タブを細かく制御 |

いずれも省略可能な引数:

- `StartURL` — 起動直後に開く URL
- `SwitchUser` — 別プロファイル名で起動（マルチインスタンス）

BiDi のみ、追加で `sessionCapabilitiesRequest` といった初期設定引数を用意しています。必要に応じて事前に `Dictionary` を組み立て、引数に渡してください。

`StartCDPMode` / `StartBiDiMode`（`CDPContext`/`WebDriverBiDiContext` を返す `Context` 版には無し）には、追加で `WebSocketMode As Boolean` 引数があります（v3.1.0〜）。`True` にすると、Pipe ではなく WebSocket 経由でローカルブラウザを起動・接続します。

```vb
Dim b As CDPBrowser
Set b = ShSetting01_StartBrowser.StartCDPMode(WebSocketMode:=True)
```

詳細は [WebSocket モードでできること](/websocket/capabilities) を参照してください。

::: tip UserForm へのブラウザ埋め込み
「UserForm の中にブラウザ画面を表示したい」場合は、この4種類とは別に [UserForm への WebView2 埋め込み](/userform/intro) を参照してください。v3.0.0 で、本キット自身が WebView2 をネイティブに CDP 制御できるようになりました。
:::

#### `sessionCapabilitiesRequest` とは

BiDi セッション確立コマンド [`session.new`](https://w3c.github.io/webdriver-bidi/#command-session-new) に渡す **Parameters** です。ブラウザ起動後・コマンド実行前に、「このセッションではどう振る舞うか」を Capabilities として宣言します（WebDriver の Desired Capabilities に相当）。

典型的な形は次のとおりです。

```vb
Dim caps As New Dictionary
Dim alwaysMatch As New Dictionary

' 例: 未処理の alert / confirm / prompt を自動で閉じず、イベントで扱う
alwaysMatch.Add "unhandledPromptBehavior", "ignore"

caps.Add "capabilities", New Dictionary
caps("capabilities").Add "alwaysMatch", alwaysMatch

Set t = ShSetting01_StartBrowser.StartBiDiModeContext( _
    StartURL:="https://example.com", _
    sessionCapabilitiesRequest:=caps)
```

- 省略時は `{ "capabilities": {} }` 相当の既定値で `session.new` します
- `StartBiDiMode` / `StartBiDiModeContext` の両方で使えます
- `reattach` でも渡せますが、**新規に BiDi-CDP Mapper が起動したときだけ** 適用されます

詳細な Capabilities 一覧は [W3C WebDriver BiDi — session.new](https://w3c.github.io/webdriver-bidi/#command-session-new) を参照してください。

## 4. 終了時は必ず quit

パイプ／マッパーを残さないよう、処理の最後でブラウザを閉じます。

```vb
' CDP
t.InheritanceCDPBrowser.quit

' BiDi
t.InheritanceWebDriverBiDiMode.quit
```

## 次へ

- [アーキテクチャ](/concepts/architecture)
- [設計思想](/concepts/design-philosophy)
- [ページ遷移](/guides/navigation)
- [要素の取得](/guides/selectors)
