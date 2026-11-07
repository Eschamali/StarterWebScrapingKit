---
description: Excel の一部の機能は WebView2 で作られています。WebView2 が Chromium をエンジンにしていることと、公式に用意された環境変数 WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS を使って、VBA から CDP で接続する仕組みを、公式ドキュメントと SDK の履歴で確かめながら解説します。
---

# からくり

## Excel の中にある WebView2

Excel の機能の一部は、WebView2 で作られています。

![Officeスクリプトのコードエディターを開いたときのプロセス一覧。WebView2 マネージャーの下に、コード エディターが並んでいる](/img/breakthrough/1.png)

Officeスクリプトのコードエディターを開いたときのプロセスの一覧です。「WebView2 マネージャー」の下に、次のような Excel の機能が並んでいます。

- WebView2: コード エディター

Office アドインについては、Microsoft のドキュメントにも、Windows では「Microsoft Edge (Chromium-based) with WebView2」で動くと書かれています。

## WebView2 は Chromium

WebView2 は Microsoft が作ったものですが、表示のエンジンには Microsoft Edge を使っています。その Edge のコアは、Google 発の Chromium です。

つまり、CDP (Chrome DevTools Protocol) も、ついでについてきちゃうということです。実際、WebView2 のデバッグ手順のドキュメントにも、「CDP のポートを開く」という手順が出てきます。

そうです、このツールは CDP を汎用的に操れます。つまり、今後 Excel が用意する新機能が WebView2 ベースなら、このツールでもすぐに操れる準備が、実質的に整ってしまうのです。

しかし当然ですが、他人のアプリが持っている WebView2 の内部メソッド（`CallDevToolsProtocolMethodForSession` など）をいじることはできません。このツールが [WebView2 経路](/stories/DebugRouteTech/webview2-tech) で使っているのは、自分で UserForm に埋め込んだ WebView2 だからこそ呼べるメソッドです。

## 公式に用意された「穴」

ところが、WebView2 について調べていくと、幸か不幸か、公式の穴が用意されています。

環境変数 `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` による、追加の起動引数の設定です。

WebView2 は、作られるときに次の環境変数を確認します（現行のドキュメントに載っているもの）。追加の起動引数は、アプリがコードで指定した値の後ろに付け足されます。

| 環境変数 | 上書きする値 |
| --- | --- |
| `WEBVIEW2_BROWSER_EXECUTABLE_FOLDER` | ランタイムのフォルダー |
| `WEBVIEW2_USER_DATA_FOLDER` | ユーザーデータフォルダー |
| `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` | 追加の起動引数 |
| `WEBVIEW2_CHANNEL_SEARCH_KIND` | チャネルを探す順番 |
| `WEBVIEW2_RELEASE_CHANNELS` | 使ってよいチャネル |

この環境変数に `--remote-debugging-port=9222` を入れておけば、アプリのコードを一切変えずに、WebView2 の CDP のポートを開けられます。同じことはレジストリ（`HKEY_CURRENT_USER\Software\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments`）でもできます。

実際、[Playwright の WebView2 のドキュメント](https://playwright.dev/docs/webview2) でも、環境変数を使う方法と、コードで `AdditionalBrowserArguments` を指定する方法の2つを挙げています。Office アドインのデバッグ手順にも、この環境変数を設定する方法が載っています。

## 環境変数が用意された経緯

この上書きの仕組みは、WebView2 の最初のプレビューにはありませんでした。NuGet に残っている SDK のヘッダー（`WebView2.idl`）を、バージョンごとに見比べると、次のように増えています。

| 日付 | SDK | 内容 |
| --- | --- | --- |
| 2019年5月 | 0.8.149 | 最初のプレビュー。上書きの記述は無い |
| 2019年6月 | 0.8.190 | レジストリ（`LoaderOverride`）で `additionalBrowserArguments` などを上書きできるようになる |
| 2019年9月 | 0.8.270 | 環境変数 `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` が加わる |
| 2020年10月 | 1.0.622.22 | 正式版。環境変数とレジストリの上書きは引き続き載っている |

0.8.190 の説明では、このレジストリの上書きを「アプリを変えずに、別のビルドの WebView2 を試すため」のものとして紹介しています。

一方で、自動テストの要望は、その後に次々と寄せられました。

- **2020年6月 #280「Test Automation - WPF + WebView2 - Hybrid Test Scenario」**: WPF と WebView2 が混ざったアプリを、Selenium などで自動テストしたいという要望です。このときの回答は「今はこのハイブリッドな自動化に良い方法がない」でした。
- **2020年12月 #715「Connecting Webview2 Office Add-in To Automation Using Selenium」**: Office アドインの中の WebView2 に、Selenium をつなぎたいという要望です。単体のアプリならつなげたが、アドインの中の WebView2 にはつなげないという相談でした。Microsoft の回答は、`--remote-debugging-port` を `AdditionalBrowserArguments` で渡せばよい、というものでした。

#715 は、まさに今回と同じ「Office が作った WebView2 を、外から自動操作したい」という相談です。ただ、アドインの作り手は、Office が WebView2 を作るときのコードには手を出せません。外から起動引数を渡せる環境変数とレジストリの仕組みが、ここで効いてきます。

今では、Microsoft の WebDriver の手順、VS Code でのデバッグ手順、Office アドインのデバッグ手順のどれもが、この仕組みを案内しています。

## 「同じユーザーなら問題ない」という考え方

外から起動引数を変えられるのは、セキュリティ的に大丈夫なのか、と思うかもしれません。

Windows には、「同じユーザーで動いているものは、もともと同じことができる」という考え方があります。Microsoft の Raymond Chen 氏は、HKEY_CURRENT_USER を書き換えるだけの「攻撃」について、こう書いています。
>それを実行できるのは、困らせたい本人（か、本人を困らせる権限のある人）だけで、ユーザーが自分の環境を困らせられることは脆弱性ではない、と。

WebView2 のドキュメントも、`WEBVIEW2_PIPE_FOR_SCRIPT_DEBUGGER` と `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` の両方を設定すると、アプリの WebView2 とその中身が、デバッガーなどの第三者のアプリに見えるようになる、と注意書きをしています。外から見えるようになることを分かったうえで、この仕組みを用意している、ということです。

## 手順

つまり、こういうことです。

1. WebView2 ベースの Excel の機能を、一旦立ち上げる
2. Excel のプロセス内の環境変数に、デバッグのポートを開く起動引数を仕込む
3. Excel の WebView2 を落として、再起動を促す
   - 偶然にも Excel は、WebView2 がクラッシュすると、一定回数までは自動で WebView2 を再起動します。これにより、`--remote-debugging-port=0` の環境変数が読まれちゃいます
4. あとは、いつも通りの CDP の儀式で制御する

2 の手順は、このツールでは1行です。

```vb
Dim WebView2Debug As New CDPCoreViaWebSocket
WebView2Debug.EnsureWebView2DebugPort = 0    'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS に --remote-debugging-port=0 を入れる
```

`EnsureWebView2DebugPort` は、Win32 API の `SetEnvironmentVariableW` で、Excel のプロセスだけに環境変数を設定します。ポートに 0 を指定すると、空いているポートが自動で選ばれ、WebView2 のユーザーデータフォルダーにある `DevToolsActivePort` にその番号が書き出されます。4 の手順では、このファイルを読んで接続します（[WebSocket 経路](/stories/DebugRouteTech/websocket-tech) と同じやり方です）。

Office アドインのデバッグ手順にあるように、システムの環境変数に設定すると、PC 上のすべての WebView2 に効いてしまいます。Excel のプロセスの中だけで設定すれば、影響は Excel が起動する WebView2 に限られます。

なお、レジストリで設定する方法もありますが、プロセス内の環境変数の方がたった1行で済むため、以降は環境変数を前提に扱います。

## 参考資料

### WebView2

* [Introduction to Microsoft Edge WebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/)（WebView2 は表示のエンジンに Microsoft Edge を使う）

* [Globals (WebView2 Win32 C++ Reference)](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/webview2-idl)（`CreateCoreWebView2EnvironmentWithOptions` が確認する環境変数とレジストリ。`WEBVIEW2_PIPE_FOR_SCRIPT_DEBUGGER` と併用すると第三者のアプリに見えるようになるという注意書き）

* [Debug WebView2 apps with Visual Studio Code](https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/debug-visual-studio-code)（環境変数かレジストリで `--remote-debugging-port=9222` を渡す。CDP のポートを開く必要がある）

* [Automate and test WebView2 apps with Microsoft Edge WebDriver](https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/webdriver)（起動済みの WebView2 に、`--remote-debugging-port` 経由で WebDriver をつなぐ「attach」の手順）

* [WebView2 | Playwright](https://playwright.dev/docs/webview2)（`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` と `AdditionalBrowserArguments` の2つの方法）

### SDK の履歴

* [Microsoft.Web.WebView2 0.8.190](https://www.nuget.org/packages/Microsoft.Web.WebView2/0.8.190)（2019年6月。`WebView2.idl` にレジストリ `LoaderOverride` の上書きが載る）

* [Microsoft.Web.WebView2 0.8.270](https://www.nuget.org/packages/Microsoft.Web.WebView2/0.8.270)（2019年9月。`WebView2.idl` に `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` が載る）

### 自動テストの要望

* [MicrosoftEdge/WebView2Feedback#280](https://github.com/MicrosoftEdge/WebView2Feedback/issues/280)（2020年6月。WPF + WebView2 の自動テスト。「今は良い方法がない」という回答）

* [MicrosoftEdge/WebView2Feedback#715](https://github.com/MicrosoftEdge/WebView2Feedback/issues/715)（2020年12月。Office アドインの WebView2 に Selenium をつなぎたいという相談。`AdditionalBrowserArguments` で `--remote-debugging-port` を渡すという回答）

### Office アドイン

* [Browsers and webview controls used by Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)（Windows では「Microsoft Edge (Chromium-based) with WebView2」）

* [Debug add-ins using developer tools for Microsoft Edge WebView2](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium)（`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` に `--auto-open-devtools-for-tabs` を入れる手順。システム上のすべての WebView2 に効く）

### セキュリティの考え方

* [It rather involved being on the other side of this airtight hatchway: Messing with somebody's registry](https://devblogs.microsoft.com/oldnewthing/20190109-00/?p=100665)（2019年1月。HKEY_CURRENT_USER を書き換えるだけでは脆弱性ではない）
