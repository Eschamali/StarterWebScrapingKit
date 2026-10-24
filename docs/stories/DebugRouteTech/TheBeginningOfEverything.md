---
description: 通信経路の話に入る前の、すべての始まり。Chrome / Edge の各チャンネルがどこに入っているかと、WebView2 をどう読み込んでいるかを、Chromium・Puppeteer・Playwright のソースと Microsoft の文書で確かめながら解説します。
---

# 技術秘話: ブラウザの探し方

本編の通信経路に入る前に、すべての始まりである「ブラウザが置いてある標準の場所」について話します。

## ブラウザのインストール場所を探す

### 対応しているブラウザ

このツールは現時点で、次の組み合わせを自動で探せます。

| ブラウザ   | 開発レベル（チャンネル）                 |
| ------ | ---------------------------- |
| Chrome | Stable / Beta / Dev / Canary |
| Edge   | Stable / Beta / Dev / Canary |

また、好きな Chromium 系ブラウザも操作できるように、ブラウザのパスを直接指定することもできます。設定ワークシートで手軽に設定できます。

### 配置のパターン

Chromium 系ブラウザ × 開発レベルの配置には、決まったパターンがあります。

**相対パス**

```
[ブラウザメーカー名]\[ブラウザ名][開発レベル]\Application\[実行ファイル名].exe
```

| ブラウザ   | Stable           | Beta                  | Dev                  | Canary               |
| ------ | ---------------- | --------------------- | -------------------- | -------------------- |
| Chrome | `Google\Chrome`  | `Google\Chrome Beta`  | `Google\Chrome Dev`  | `Google\Chrome SxS`  |
| Edge   | `Microsoft\Edge` | `Microsoft\Edge Beta` | `Microsoft\Edge Dev` | `Microsoft\Edge SxS` |

* 実行ファイル名は、Chrome が `chrome.exe`、Edge が `msedge.exe` です。フォルダ名の \[ブラウザ名] と、実行ファイル名は必ずしも一致しません

* Stable には、開発レベルの名前が付きません。Canary は「SxS」になっています

SxS は「Side-by-Side」のことで、通常版と並べて同じ PC に入れられる版、という意味です。Chromium のソース（`chrome/install_static/google_chrome_install_modes.h`）では、Beta・Dev・Canary の3つを「the side-by-side … channel」と書き、そのうち Canary の名前にだけ「 SxS」を付けています。なお、SxS が Side-by-Side の略だと明記した一文は、ソースの中には見つかりませんでした。

**親フォルダ**

親フォルダは、インストールするときの**権限**で2つのルートに分かれます。

| ルート    | 入る場所                                       | いつこうなるか                  |
| ------ | ------------------------------------------ | ------------------------ |
| マシン単位  | `%ProgramFiles%` または `%ProgramFiles(x86)%` | 管理者権限で入れたとき              |
| ユーザー単位 | `%LOCALAPPDATA%`                           | 管理者権限が無いとき。Canary は常にこちら |

製品版を配布ページから入れた場合、開発レベルとの関係は次のとおりです。

* **Chrome の Stable〜Dev**：管理者権限があればマシン単位、無ければユーザー単位です。Google の配布ページ（`www.google.com/chrome/`）の JavaScript は、Windows 向けのインストーラーに `needsadmin=prefers` を付けています。Google Update（Omaha）のソース（`omaha/common/const_goopdate.h`）では、`prefers` は「権限があればマシン単位、無ければユーザー単位」と定義されています。管理者権限が無い場合や、UAC で「いいえ」を選んだ場合は、`%LOCALAPPDATA%` に入ります

* **Edge の Stable〜Dev**：マシン単位です。Microsoft の文書（WebView2 の配布の説明）に、Canary 以外の Edge にはマシン単位の Edge Updater が付いてくる、とあります。手元の PC の Edge Stable も、`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe` にありました。ただし、会社のポリシー（Edge Update の「Force Installs (Per-User)」）でユーザー単位にもできます

* **Canary（Chrome・Edge）**：常にユーザー単位で、`%LOCALAPPDATA%` に入ります。Chrome は配布ページが Canary だけ `needsadmin=false` を付けており、Chromium のソース（`google_chrome_install_modes.h`）でも、Canary はマシン単位に対応していません（`supports_system_level = false`）。Edge も、マシン単位の Edge Updater が付いてこないのは Canary だけです

マシン単位のときに `%ProgramFiles%` と `%ProgramFiles(x86)%` のどちらになるかは、Chrome では入れた時期で変わります。

* 2020年6月の Chromium の変更（「Install 64-bit browser versions under "C:\Program Files" by default」）以降に新しく入れた 64bit 版は、`C:\Program Files`

* それより前に `C:\Program Files (x86)` に入れたものは、そこに入ったまま更新され続ける

ここまで押さえれば、自動で探すロジックは一通り作れます。ユーザー自身がインストール場所を変えない限り、このロジックで見つかります。

### このツールが探す順番

Stable / Beta / Dev は、次の順に探し、最初に見つかった場所を使います。Canary は、ユーザー単位（`%LOCALAPPDATA%`）だけを見ます（`CDPHost.cls` の `getBrowserPath`）。

1. `%ProgramFiles(x86)%`（マシン単位。以前から入れている Chrome や、Edge）
2. `%ProgramW6432%`（マシン単位。最近入れた 64bit 版の Chrome）
3. `%LOCALAPPDATA%`（ユーザー単位。管理者権限が無かった、UAC で「いいえ」を選んだ、Edge のポリシーでユーザー単位にした、などの場合）

2番目に `%ProgramFiles%` ではなく `%ProgramW6432%` を使っているのは、32bit 版の Excel から `%ProgramFiles%` を読むと、`C:\Program Files (x86)` が返ってくるためです。`%ProgramW6432%` なら、Excel が 32bit でも 64bit でも `C:\Program Files` が返ります。この変数が取れない環境では、`%ProgramFiles%` を使います。

::: info 自分でビルドした Chromium の場合
Chromium のインストーラー（`mini_installer.exe`）を直接実行した場合は、既定がユーザー単位（`%LOCALAPPDATA%`）で、`--system-level` を付けたときだけマシン単位になります（`chrome/installer/mini_installer/README.md`）。このツールの自動探索は製品版を対象にしているので、汎用の Chromium は設定ワークシートでパスを指定してください。
:::

レジストリを追えば、ユーザーが変えたインストール場所も特定できるかもしれません。ただ、このツールはそこまでは追いません。面倒なので…。

Puppeteer や Playwright も、そこまでは追っていません。

* Puppeteer（`packages/browsers/src/browser-data/chrome.ts`）：`PROGRAMFILES`、`ProgramW6432`、`ProgramFiles(x86)`、`LOCALAPPDATA` の環境変数と、`C:\Program Files` などの予備の候補に、決まった相対パスをつなげて探します

* Playwright（`packages/playwright-core/src/server/registry/index.ts`）：`LOCALAPPDATA`、`PROGRAMFILES`、`PROGRAMFILES(X86)`、`HOMEDRIVE` の `Program Files` に、決まった相対パスをつなげて探します

どちらのソースにも、ブラウザの場所をレジストリから読む処理はありません。インストール場所を変えた場合は、このツールと同じく、パスを直接指定する作りです。このツールでは、ここは先ほどの設定ワークシートに任せています。

## WebView2 について

このツールには、WebView2 を制御する仕組みも入っています。では、WebView2 はどうやって探しているのでしょうか。

実はこれは、`WebView2Loader.dll` に任せています。

「あれ、このツールって外部依存禁止の縛りじゃなかったっけ？」
はい、そうです。それは守っています。少なくとも、今の基準では。

実は、最新の Office（特に Microsoft 365）には、この `WebView2Loader.dll` が最初から入っています。手元の PC では、`C:\Program Files\Microsoft Office\root\Office16\WebView2Loader.dll` にありました（ファイルの製品名は「Microsoft Edge Embedded Browser WebView Loader」、バージョンは 1.0.3485.44）。

Microsoft の文書でも、Microsoft 365 Apps の Version 2101 以降には、2021年4月から WebView2 Runtime を入れ始めたこと、Office アドインが WebView2 を使っていることが書かれています。一方で、Office に `WebView2Loader.dll` を同梱していると書いた文書は、見つかりませんでした。

VBA は、Office のインストールのついでに入った `WebView2Loader.dll` を使っているだけです。なので、`LoadLibrary` で DLL の場所を指定する、といった外部 DLL を読み込むための手順は要りません。普段の Win32 API と同じ感覚で、宣言を書くだけで動きます。

このツールも、`Declare` で `Lib "WebView2Loader"` とパスなしで書いているだけです（`WebView2Host.cls`）。それでも動くのは、Windows が DLL を探す順番のおかげです（Microsoft Learn「Dynamic-link library search order」）。

1. すでに読み込まれている DLL の一覧
2. アプリケーションの実行ファイルがあるフォルダ

`WebView2Loader.dll` は `EXCEL.EXE` と同じ `Office16` フォルダにあるので、パスを書かなくても見つかります。呼び出した後が本当の地獄なのですが、それは少し後のお話で…。

…つまり逆に言えば、`WebView2Loader.dll` が入っていない古い Office では動きません。古い Office の頃には、まだ WebView2 という存在すら知られていなかったので。なお、どのバージョンの Office から入っているかは、手元に古い Office が無いため確かめられていません。

### WebView2Loader.dll すら使わない方法

これはまだ試していませんが、どうやら `WebView2Loader.dll` すら使わない POC が、ネットにまさかの形で転がっています。

* [WebView2をC言語から利用する（WebView2Loader.dll不使用）（Qiita）](https://qiita.com/tkzn_io/items/cda871df3110a24cc3ae)

* [takazin11de/my\_webview2（GitHub）](https://github.com/takazin11de/my_webview2/blob/main/my_webview2.c)

簡潔に言うと、次のことをしています。

1. レジストリから、WebView2 の置き場所を読む（`HKEY_LOCAL_MACHINE\Software\Microsoft\EdgeUpdate\ClientState\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}` の `EBWebView` という値。手元の 64bit の Windows では、32bit 用の `WOW6432Node` の下にありました。OpenWebView2Loader も、32bit 側を見る指定 `KEY_WOW64_32KEY` を付けて開いています）
2. その場所の `EBWebView\x64\EmbeddedBrowserWebView.dll` を読み込む
3. その中の `CreateWebViewEnvironmentWithOptionsInternal` という関数を呼ぶ

そうです。この `EmbeddedBrowserWebView.dll` が、真の WebView2 のコアです。`WebView2Loader.dll` は名前からお察しの通り、WebView2 を探して読み込むだけの、言ってしまえば簡易なヘルパー DLL です。Qiita の記事でも、ローダー DLL がしていることは「レジストリを読んでWebView2のインストールフォルダを特定し、WebView2（COMサーバー）を起動するぐらい」と書かれています。

置き場所は、ブラウザの探し方とは少し違います。

| チャンネル                      | 置き場所の例                                                                                                         |
| -------------------------- | -------------------------------------------------------------------------------------------------------------- |
| WebView2 Runtime（Stable）   | `%ProgramFiles(x86)%\Microsoft\EdgeWebView\Application\[バージョン]\EBWebView\[アーキテクチャ]\EmbeddedBrowserWebView.dll` |
| Edge の Beta / Dev / Canary | 各ブラウザの `Application\[バージョン]\EBWebView\[アーキテクチャ]\`                                                              |

* Stable は、Edge ブラウザ本体ではなく、別に入っている **WebView2 Runtime**（`EdgeWebView`）を使います。手元の PC のレジストリでも、上の `{F3017226-…}` は「Microsoft Edge WebView2 Runtime」で、`EBWebView` の値は `EdgeWebView` のフォルダを指していました。Edge Stable のフォルダにも `EBWebView` はありますが、Edge Stable のレジストリには `EBWebView` の値がありませんでした

* WebView2Loader の再実装（[OpenWebView2Loader](https://github.com/jchv/OpenWebView2Loader)）も、同じ順番で探しています。探すチャンネルは、Stable（WebView2 Runtime）、Beta、Dev、Canary です

* \[アーキテクチャ] は、`x64`、`x86`、`arm64` のどれかです。呼び出す側（Excel）のビット数に合わせます

見ての通り、一般的なブラウザの探し方よりも、次の情報が追加で必要になります。

* バージョン（レジストリの値に含まれている）

* アーキテクチャ（Excel のビット数）

これを自動で特定できれば、もしかしたら古い Office でも動くかもしれない…そんなロマンのあることが起こるかもしれないのです。

しかし、`CreateWebViewEnvironmentWithOptionsInternal` は、あくまでも非公開の API です。OpenWebView2Loader のソースにも「This routine is not documented!」（この関数は文書化されていない）とあり、その README も、たいていの用途では公式の WebView2Loader を使うべきだと書いています。Edge の更新で、突然動かなくなるリスクが増えます。

…が、時間があったらやってみる価値はあるかもしれません。

手に入るのは、「古い Office でも WebView2 が動く」。ただそれだけかもしれません。サポートが終わった製品まで気にしてしまうなんて、ボクはどうしてしまったんだろうね。

さて、ブラウザの探し方はこの辺にして、次は本編の各通信経路について話します。最初は [Pipe 経路](/stories/DebugRouteTech/pipe-tech)です。

## 参考資料

* [chrome/installer/mini\_installer/README.md](https://source.chromium.org/chromium/chromium/src/+/main:chrome/installer/mini_installer/README.md)（Chromium のインストーラーを直接実行したときの、マシン単位とユーザー単位のインストール場所）

* [google/omaha omaha/common/const\_goopdate.h](https://github.com/google/omaha/blob/main/omaha/common/const_goopdate.h)（`needsadmin` の `prefers` の定義）

* [Chromium の updater の仕様（functional\_spec.md）](https://chromium.googlesource.com/chromium/src/+/main/docs/updater/functional_spec.md)（`needsadmin=true` / `prefers` のときに管理者として起動し直す動き）

* [Reland "Install 64-bit browser versions under "C:\Program Files" by default"](https://github.com/chromium/chromium/commit/bf476e406173d06ab3b63453ca63e0fa485d43e8)（2020年6月）

* [Distribute your app and the WebView2 Runtime（Microsoft Learn）](https://learn.microsoft.com/ja-jp/microsoft-edge/webview2/concepts/distribution)（Canary 以外の Edge にはマシン単位の Edge Updater が付いてくる）

* [Microsoft Edge Update policies（Microsoft Learn）](https://learn.microsoft.com/ja-jp/deployedge/microsoft-edge-update-policies)（マシン単位とユーザー単位のポリシー）

* [chrome/install\_static/google\_chrome\_install\_modes.h](https://source.chromium.org/chromium/chromium/src/+/main:chrome/install_static/google_chrome_install_modes.h)（side-by-side のチャンネル、` SxS`、Canary はマシン単位に非対応）

* [Puppeteer packages/browsers/src/browser-data/chrome.ts](https://github.com/puppeteer/puppeteer/blob/main/packages/browsers/src/browser-data/chrome.ts)

* [Playwright packages/playwright-core/src/server/registry/index.ts](https://github.com/microsoft/playwright/blob/main/packages/playwright-core/src/server/registry/index.ts)

* [Dynamic-link library search order（Microsoft Learn）](https://learn.microsoft.com/ja-jp/windows/win32/dlls/dynamic-link-library-search-order)

* [Microsoft 365 Apps と WebView2 Runtime（Microsoft Learn）](https://learn.microsoft.com/ja-jp/microsoft-365-apps/deploy/webview2-install)

* [Office アドインで使われるブラウザー（Microsoft Learn）](https://learn.microsoft.com/ja-jp/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)

* [OpenWebView2Loader](https://github.com/jchv/OpenWebView2Loader)（WebView2Loader の再実装）

* [WebView2をC言語から利用する（WebView2Loader.dll不使用）（Qiita）](https://qiita.com/tkzn_io/items/cda871df3110a24cc3ae) / [takazin11de/my\_webview2](https://github.com/takazin11de/my_webview2)
