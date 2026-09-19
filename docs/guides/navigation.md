---
description: CDP / BiDi でのページ遷移（navigate）と読み込み待ち。ReadyState による待機条件の指定方法をコード例付きで解説します。
---

# ページ遷移

タブ（コンテキスト）に対して URL を開き、読み込み完了を待ちます。

## 基本

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

t.navigate "https://kemono-friends.jp/"          ' 既定: 読み込み完了まで待つ
t.navigate "https://kemono-friends.jp/introduction/", WaitMode:=isComplete
t.Wait                                    ' 現在ページの完了待ち（ポーリング版）

t.ThisCDPBrowser.quit
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

t.navigate "https://kemono-friends-20170110.jp/"

t.ThisWebDriverBiDiMode.quit
```

:::

起動時 URL を渡すこともできます。

::: code-group
```vb [CDP]
Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://kemono-friends.jp/")
```

```vb [BiDi]
Set t = ShSetting01_StartBrowser.StartBiDiModeContext("https://kemono-friends-20170110.jp/")
```

:::

## 起動時のレースコンディション対策（v3.2.0〜）

`--app="<URL>"` でブラウザを起動する場合、以前は起動直後にCDP/BiDiコマンドを送ると、ブラウザ側の準備が間に合わず稀に失敗することがありました（例: 起動直後の`Page.navigate`が最初の1回だけ失敗する）。

v3.2.0からは、**Context（タブ）単位で起動する場合**、実際のURLへ向かう前に一度「起動スプラッシュ画面」（`data:text/html,...`の自己完結HTML、ライト/ダーク自動対応のローディング表示）を経由するようになりました。ライブラリは確実にこのスプラッシュタブへ接続し、`Page.enable`/`Runtime.enable`を有効化してから、あらためて本来のURLへ`navigate`します。これにより「起動直後の最初のコマンドだけ失敗する」という問題が解消されています。

![起動直後に一瞬だけ表示される、ダークテーマ対応の起動スプラッシュ画面。「Starting CDP control from VBA...」というタイトルとスピナー、遷移先URLが表示されている](/img/起動スプラッシュ画面.png)

*▲ 実際の起動スプラッシュ画面。この裏で `Page.enable` / `Runtime.enable` が有効化されたあと、指定したURLへ本遷移します*

```vb
Sub 起動シーケンスのイメージ()
    '1. --app="data:text/html,...(スプラッシュ画面)" でブラウザを起動
    '2. スプラッシュタブへ確実に接続（Page.enable / Runtime.enable も有効化）
    '3. ここで初めて、指定したURLへ navigate する
    Dim t As CDPContext
    Set t = ShSetting01_StartBrowser.StartCDPModeContext("https://example.com")
    ' ↑ この行が完了した時点で、上記1〜3は既に終わっている
End Sub
```

::: tip Context単位のみで発生。ブラウザ単位は対象外
このスプラッシュ画面が挟まるのは `StartCDPModeContext` / `StartBiDiModeContext`（およびそれらが内部で呼ぶ `CDPContext.StartAndConnectTab` / `WebDriverBiDiContext.StartBiDiModeAndConnectTab`）のみです。`StartCDPMode` / `StartBiDiMode`（`CDPBrowser` / `WebDriverBiDiMode` を返すブラウザ単位の起動）は、以前と同じく `--app` に直接URLを渡します。
:::

## `Wait` と `WaitEvents`（v3.2.0〜）

読み込み完了を待つ方法が2種類あります。

| | `Wait` | `WaitEvents` |
| --- | --- | --- |
| 方式 | `document.readyState` を**ポーリング** | CDP/BiDiの読み込みイベントを**受動的に待つ** |
| 使い所 | `getTab` / `newTab` で取得した、いつ遷移したか分からないタブ | `navigate` や、直前に自分で起こした遷移操作の完了待ち |
| 事前準備 | 不要 | 遷移を起こす**前**に [`ResetWaitState`](/api/cdp/CDPContext#resetwaitstate) が必要 |

`navigate` は内部で `WaitEvents` を使っています。クリックなど、`navigate` を経由しない遷移（フォーム送信やSPA内リンクなど）の完了を待ちたい場合は、次のように自分で呼び出します。

::: code-group
```vb [CDP]
t.ResetWaitState
t.getElementByQuery("a.next").click
t.WaitEvents WaitMode:=isComplete
```

```vb [BiDi]
t.ResetWaitState
t.UpgradeBiDiPlus.getElementByQuery("a.next").click   ' 要素クリックはCDP側のAPIを使う例
t.WaitEvents WaitMode:=isComplete
```
:::

::: warning `ResetWaitState` を忘れると…
`WaitEvents` は「状態が変化したこと」をイベントで検知する仕組みです。`ResetWaitState` を呼ばずに使うと、前回までに到達済みの状態がそのまま返ってしまうことがあります。**いつ遷移したか分からないタブ**（新しく開いた直後のタブなど）では、素直に `Wait`（ポーリング版）を使ってください。
:::

## `navigate` の引数拡張（CDP、v3.2.0〜）

CDPの `navigate` は、[`Page.navigate`](https://chromedevtools.github.io/devtools-protocol/#/Page.navigate) の全パラメータ（`referrer` / `transitionType` / `referrerPolicy` / `frameId`）に対応しました。`referrer` がないとアクセスできない高度なサイトなどに使えます。

```vb
' referrer を偽装しないとアクセスできないサイト向け
t.navigate "https://example.com/protected", referrerURL:="https://example.com/"

' 待たずに次へ進みたい場合
t.navigate "https://example.com", WaitMode:=Nowait
```

::: warning 破壊的変更：第2引数の意味が変わりました
以前は `navigate(strURL, till As ReadyState)` でしたが、v3.2.0で第2引数は `referrerURL As String` になりました。`t.navigate url, isInteractive` のように**第2引数を位置引数で待機条件として渡していたコードは動作が変わります**。`t.navigate url, WaitMode:=isInteractive` のように名前付き引数へ書き換えてください。詳細は [`CDPContext.navigate`](/api/cdp/CDPContext#navigate) を参照。BiDi側の `navigate` は今回変更されていません。
:::

## ReadyState（CDP / BiDi 共通）

`navigate` / `Wait` / `WaitEvents` / 一部の要素操作は `ReadyState` で待機条件を選べます。よく使うのは `isComplete`（ドキュメント完了）です。v3.2.0で「待たない」ことを明示する `Nowait` が追加されました（`navigate`/`WaitEvents`の`WaitMode`に渡すと、イベントを待たずに戻ります）。

## 関連 API

- [`CDPContext.navigate`](/api/cdp/CDPContext#navigate) / [`WaitEvents`](/api/cdp/CDPContext#waitevents) / [`ReadyState`](/api/cdp/CDPContext#readystate)
- [`WebDriverBiDiContext.navigate`](/api/bidi/WebDriverBiDiContext#navigate) / [`WaitEvents`](/api/bidi/WebDriverBiDiContext#waitevents)
- [マルチタブ](/guides/multi-tab)
