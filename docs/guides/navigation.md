# ページ遷移

タブ（コンテキスト）に対して URL を開き、読み込み完了を待ちます。

## 基本

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext

t.navigate "https://kemono-friends.jp/"          ' 既定: 読み込み完了まで待つ
t.navigate "https://kemono-friends.jp/introduction/", isComplete
t.wait                                    ' 現在ページの完了待ち

t.InheritanceCDPBrowser.quit
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext

t.navigate "https://kemono-friends-20170110.jp/"

t.InheritanceWebDriverBiDiMode.quit
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


## ReadyState（CDP）

`navigate` / `wait` / 一部の要素操作は `ReadyState` で待機条件を選べます（クラス定義の列挙を参照）。よく使うのは `isComplete`（ドキュメント完了）です。


## 関連 API

- [`CDPContext.navigate`](/api/cdp/CDPContext#navigate)
- [`WebDriverBiDiContext.navigate`](/api/bidi/WebDriverBiDiContext#navigate)
- [マルチタブ](/guides/multi-tab)
