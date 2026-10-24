---
description: WebDriver なしでWebDriverBiDiが扱えるとわかった経緯。公式ドライバーに頼らない BiDi 対応の登場秘話です。
---

# なぜ外部依存なしで、ChromiumにWebDriverBiDiが扱えるとわかったのか

まずは映像で、ひと通り見てください。

<PromoMovie src="/promo/ReiwaIE_BiDiStory.html" title="令和のIE 技術秘話 ― BiDi Story" />

![海外フォーラムでの議論](/img/story/1.png)
すべては[この議論](https://github.com/GCuser99/SeleniumVBA/discussions/164#discussioncomment-15988897)から始まりました。

> Well, I think we had this discussion earlier regarding BiDi. CDP is not futureproof as BiDi is.
> And when it comes to companies not allowing local software installations: if Chrome is not being allowed to install, then CDP is worthless anyway. So even in a scenario of restricted environments I assume CDP support not to be a breakthrough.
さて、BiDiについては、以前にもこの話をしたと思います。CDPには、BiDiほどの将来性がありません。
そして、企業がローカルへのソフトウェアのインストールを許可しない場合についていえば、Chromeのインストールが許可されていないなら、CDPはどのみち無意味です。ですので、制限のある環境を想定した場合でも、CDPへの対応は画期的なものではないと考えます。

この頃になると、CDPに対する知識は随分と溜まってきました。
なので、すぐに気付きました。半分合っていて、半分間違ってる、と。

CDPは、Chromeブランドではなく、Chromiumブランドです。CDP の[公式ページ](https://chromedevtools.github.io/devtools-protocol/)も、対象を「Chromium、Chrome、そのほかの Blink 系ブラウザ」としています。よって、業務PCにプリインストールされているEdgeも、CDPに対応していることになります。Microsoft も、Edge の DevTools Protocol は Chrome DevTools Protocol と API が一致すると[説明しています](https://learn.microsoft.com/ja-jp/microsoft-edge/devtools/protocol)。
今思えば、あのMicrosoftのEdgeごり押し戦法

・世間からMicrosoftに批判ラッシュがあっても
・大手ニュースから批判されても
・アンインストールできる仕組みはあるのに、アンインストールさせない罠によるバッシング

それでもMicrosoftはWindowsのEdgeを推し進めた。
結果的に、ブック1枚でのブラウザ自動化が、IEの頃から(偶然にも？)受け継がれたのです。
これは紛れもない事実です。

でも...**「将来性はBiDiが上」**
これも、ある意味、事実です。

その言葉が頭から離れなかったのです。

こうして、私は決めたのです。
>絶対に、外部依存なしで、ChromiumをWebDriverBiDiで扱える方法を探してやる

と。

本題に入る前に、WebDriverBiDiに至るまでの歴史を、少しサクッと述べようと思います。

## WebDriverの歴史

WebDriverBiDiの立ち位置は、CDPと同じです。決められた形式に従ってJSONの文字列を書き、ブラウザに送ります。違いは、誰がルールを決めているかです。
CDP:Google（Chromium）
WebDriverBiDi:W3C

W3Cは簡単に言うと、大手のブラウザメーカーたちが集まって共通のコマンド名でうちらのブラウザを操る世界標準を作るグループみたいなもんです。

全ての始まりは、GoogleでもMicrosoftでもなく、Seleniumです。2004年に ThoughtWorks の Jason Huggins が作ったテストツールです（[Selenium 公式の歴史](https://www.selenium.dev/history/)）。
そうです。VBAでは標準となってる`SeleniumVBA`も、Seleniumというプロジェクトの上に成り立っています。Selenium の WebDriver を使い、ドライバーが無ければ自動で落としてくる作りです（[SeleniumVBA の README](https://github.com/GCuser99/SeleniumVBA)）。

最初の自動化はどうしてたのかというと...
テスト用の JavaScript（Selenium Core）を、テストしたいサイトと同じサーバーに置き、ブラウザの中で動かしていました。ブラウザには、JavaScript からはページを配ったサーバーにしかアクセスさせない決まり（同一オリジンポリシー）があるからです（[Selenium WebDriver の作者 Simon Stewart による解説](https://aosabook.org/en/v1/selenium.html)）。

でも、テストのたびに相手のサーバーへ置くのは現実的ではありません。そこで登場した Selenium RC がやったのは...
なんと！ネットワークの書き換えです。

要するに、間にプロキシサーバーを挟み、そこで受け取ったHTMLにテストツールのJSを仕込むという、今見たらもはやウイルスチックなことをしていました（[Selenium RC の公式ドキュメント](https://www.selenium.dev/documentation/legacy/selenium_1/)でいう「Proxy Injection」）。でも当時は、これしか選択肢がなかったのです。
しかし当然、この仕組みには限界があります。
あくまでも主体はJavaScriptなので、JavaScriptのルール違反になるようなことをしようとすると、動きません（ファイルのアップロードや、ドメインをまたぐ移動が例に[挙がっています](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)）。
でも、そのルール違反的なことをしないと、テストもままならないのも事実...

そして時が経ち、ようやくおなじみのWebDriverが登場しました...いや厳密には、「WebDriver」という名前の製品が登場しました。最初のコードが公開されたのは、2007年の初めです（[Simon Stewart による解説](https://aosabook.org/en/v1/selenium.html)）。

でも内部ロジックとしては今と同じように
・ブラウザごとの操作内容を隠ぺい/共通化
しています。

例えば、このツールのIntroとして出してるIEならCOM
別ブランチ(ex-Firefox)で密かに対応してるFirefoxなら、当時は拡張機能（XPCOM）
今主流のChromeなら、当時は拡張機能＋NPAPI プラグイン

IE と Firefox のやり方は、作者の Simon Stewart 自身が[ブログ](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)と[解説](https://aosabook.org/en/v1/selenium.html)で説明しています。IE は MSHTML と ShDocVw の COM インターフェースを使い、Firefox は XPCOM を使う拡張機能です。Chrome については、Selenium 2.0 alpha 1（2009年）の Chrome ドライバーが、「WebDriver extension for Chrome」という拡張機能に [NPAPI のプラグイン `npchromedriver.dll`](https://github.com/SeleniumHQ/selenium/blob/05ceaf552081cba0035862ea0483c08b2692338e/chrome/src/extension/manifest-win.json#L4-L17) を積んだ作りでした。

当然、こんなバラバラな相手を1つ1つ学習してたらキリがありません。
それを、WebDriverという製品が解決してくれました。

そしてSeleniumは気付きました
・うちのJavaScriptハックよりもこのWebDriverのほうがええやん

こうしていい感じに統合されて、ここでようやく、規格としてのWebDriverが登場しました。統合の[発表](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)が2009年、[Selenium 2.0 としての公開](https://www.selenium.dev/blog/2011/selenium-2-0/)が2011年、[W3C 勧告](https://www.w3.org/TR/2018/REC-webdriver1-20180605/)が2018年です。
複数のブラウザが相手なのに、覚えるAPIはWebDriverという、巨大ではあるものの1種類だけ。
当時は画期的でした。

しかし、Webは異常な速さで進化しました。
ちょっと思い出してみてください。

昔、ブラウザでゲームをするには、
・Flash Player
・Java

のような、今でいう拡張機能を入れる必要がありました。
「これを使うには、Flash Playerをインストール...」
一体何回見たことやら...

また、当時のYouTubeも、実は動画の部分をFlash Playerに依存していました。YouTube が HTML5 のプレーヤーを既定にしたのは、[2015年1月](https://web.archive.org/web/20160101205917/http://youtube-eng.blogspot.com/2015/01/youtube-now-defaults-to-html5_27.html)です。

でも今のブラウザはどうです？
・URLを共有するだけで動画もゲームもすぐにできる/見れる

理由は単純です。それらをつかさどる機能が、ブラウザに次々と丸ごと標準搭載されたからです。
・WebGL API
・Web Audio API
・WebAssembly
・Canvas
...AndMore

こうした技術が短い間に次々と実装され、リンクを共有するだけで素晴らしい体験ができるようになりました。

...しかし、この進化に追いつけてないのが1つだけありました。
はい、WebDriverです。

WebDriverの基本は、コマンドを送り、応答を待って結果を受け取る、という同期のスタイルです（技術的にいうと、HTTPによる同期のやり取り）。
Chrome チームも、WebDriver Classic は HTTP でやり取りするので、ブラウザのイベントを購読して受け取るのが難しいと[書いています](https://developer.chrome.com/blog/test-automation-evolution)。
しかし、さきほどの進化の影響で、この仕組みに限界が出てきました。
・ページは読み込み終わったのに、表示されている内容が書き換わる
　→検索ボタンを押すと、ページを丸ごと読み直すのではなく、一部だけ読み直す

・inputタグの動的生成への対応
　→初期状態にはなく、ボタンを押して初めて現れるタイプ。初期状態で探しても見つからない

・ネットワークの監視
　→非同期処理で届くデータの内容に応じた処理

いずれも、SPA（Single Page Application）のサイトで起きる「内部で起こった出来事」です。SPAは、技術が発展したからこそ登場したサイトです。

しかし実は、これら3つとも解決できる手段は、あったっちゃありました。
はい、このツールの主題であるCDPです。CDPでは、コンソールのメッセージを拾ったり、ネットワークのリクエストを横取りしたりといった低いレベルの操作ができ、WebSocket で双方向にやり取りできます（[同じ解説](https://developer.chrome.com/blog/test-automation-evolution)）。

CDPは、元々はデバッグ用のプロトコルです。今の CDP につながるのは、WebKit の Web Inspector（開発者ツール）をブラウザの外から使えるようにするため、2011年に[発表](https://web.archive.org/web/20110513094715/http://www.webkit.org:80/blog/1620/webkit-remote-debugging/)された「WebKit Remote Debugging Protocol」です。発表の記事には `Chromium --remote-debugging-port=9222` で起動する例が載っていて、v1.0 は[2012年](https://web.archive.org/web/20120507025807/http://www.webkit.org/blog/1875/announcing-remote-debugging-protocol-v1-0/)に出ました。意外にも、WebDriver（2007年）より後発です。
いまの CDP の定義は、[Chromium のソースの中](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:third_party/blink/public/devtools_protocol/)にあります。当然、Webの進化とともに、CDPも進化していきました。
そして進化していく中で、気づいてしまいました。
>あれ？これブラウザ自動化にも使えるじゃん

これが、あのPuppeteerの登場です。Puppeteer は、Chrome チームが作った Node のライブラリです（[Headless Chrome の解説](https://developer.chrome.com/blog/headless-chrome/)）。CDP について Chrome チーム自身も、「もとは DevTools とデバッグのために作られ、Puppeteer が自動化に取り入れた」と[書いています](https://developer.chrome.com/blog/test-automation-evolution)。

しかし、名前のとおり、あくまでもChromium(Chrome/Edge)向けのプロトコルです。Firefox も一時期、Puppeteer のために CDP の一部を実装していましたが、恒久的な手段ではないとして WebDriver BiDi に移っています（[Chrome チームの記事](https://developer.chrome.com/blog/firefox-support-in-puppeteer-with-webdriver-bidi)）。Selenium も、4.29.0 で [Firefox 向けの CDP 対応を外しました](https://www.selenium.dev/blog/2025/remove-cdp-firefox/)。
このままでは、またあの
・各ブラウザを相手に1つ1つ学習してたら鬱になる

が戻ってしまいます。

こうして、現在のWebDriver BiDiの制定に至りました。2019年の W3C の会議（TPAC）で[設計が話し合われ](https://github.com/w3c/webdriver-bidi/blob/f922e6af669ead51496db85f37813f39438ef171/explainer.md)、[仕様](https://w3c.github.io/webdriver-bidi/)は、WebDriver に双方向の通信を足して、ブラウザからイベントを流せるようにするものとして書かれています。

## ChromiumをWebDriverBiDiとして話す方法

さて、お待たせしました。ようやく、本題です。
大まかな歴史をつかんだところで、このツールはどうやって、ChromiumをBiDiとして喋っているのか...

### WebDriver.exeルート

とその前に、一般的に、WebDriver.exe経由でChromiumをWebDriverBiDiで扱うにはどうするのか？
ざっくりした流れは、こうです。
1. DLしたWebDriverを、「--port=9515」のようにポートを指定して起動する
2. そのポートの`/session`へ、capabilitiesに「webSocketUrl: true」を付けて New Session を送る
3. すると ChromeDriver は内部で、mapper（後述）用の非表示タブを立ち上げ、そこへ`session.new`を送ってWebDriverBiDiを有効にする。その応答の`session.NewResult`は、ChromeDriver の中で受け取る
4. New Session のレスポンスの capabilities の`webSocketUrl`に、`ws://<ホスト>/session/<セッションID>`が入って返るので、そこへ WebSocket で接続する
5. あとはBiDiのルールに沿って、WebSocket上でコマンドをやり取りする

2〜4は、W3C の仕様と ChromeDriver のソースで確かめられます。[仕様](https://w3c.github.io/webdriver-bidi/#establishing)には、WebSocket URL の capability を true で頼むと双方向の接続になる、とあります。ChromeDriver のソースで当たるのは、[非表示タブを作る所](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=411-418)、[`session.new` を送る所](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=464-475)、[`webSocketUrl` を返す所](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=288-291)です。なお、仕様の`session.new`は、HTTP を使わずに BiDi だけでセッションを作るためのコマンドです（[仕様](https://w3c.github.io/webdriver-bidi/#command-session-new)）。

WebDriver.exeでBiDiを扱う詳しい流れは、[こちらの記事](https://note.com/sele_chan/n/n2b21c7c26ef8)で確認できます。

### このツール

では、このツールの場合はどうなのか...
大きく2フェーズあります。
こうなります。

<準備編>
1. npm等で`mapperTab.js`を入手（[chromium-bidi](https://www.npmjs.com/package/chromium-bidi) のパッケージに入っています）
2. 中身をExcel/Accessに保存
　→Excelの場合、セルに入る文字数の上限が[32,767 文字](https://support.microsoft.com/ja-jp/office/excel-worksheet-and-workbook-limits-c13798fc-f1d5-4e73-998d-c02268bc65ae)なので、百万字近くある`mapperTab.js`（v17.0.2 で約85万字）を、その文字数ずつ分割して保存してます。
　　なお、分割したかたまりが偶然 `=` で始まると、Excel が数式として読もうとして「メモリが不足しています」という、つじつまの合わないエラーになることがありました。今は、各かたまりの先頭に `'` を付けて、文字列として入れています
　　Accessには、1レコードに収められる「長いテキスト」型（[最大 1 GB](https://support.microsoft.com/ja-jp/access/data-types-for-access-desktop-databases)）があるので、特に苦労なく入れられた

<実際に動かす>
1. BiDiとして使用を開始する`Sub`を呼び、Chromiumのデバッグ(Port/Pipe)に接続
2. 準備編で保存した`mapperTab.js`を変数に確保する
　→Excelの場合は、分割したやつをJoinする。AccessはシンプルにSelect文

3. 新しいタブ("Target.createTarget")を非表示("hidden:=True")で開く　※必要に応じて表示する設定も可能。以降は、3で開いたタブに対する操作となる
4. `Target.exposeDevToolsProtocol`で、JavaScriptからCDPを操作する権限を与える
5. `Runtime.addBinding`で、BiDiのレスポンスを受け取る名称（`sendBidiResponse`）を設定。任意でデバッグ用の名称（`sendDebugMessage`）も設定します
6. `Runtime.evaluate`で、`mapperTab.js`の内容を実行する
7. 6で定義された初期化用の関数`runMapperInstance`に、3で開いたタブのIDを渡して初期化
8. BiDiコマンドを`Runtime.callFunctionOn`の引数に載せて、タブの中の`window.onBidiMessage`を呼ぶ
9. `Runtime.bindingCalled`を監視し、BiDiの結果なら、RaiseEventでBiDi用のClassへ伝える

::: details 出典：各手順がソースのどこにあるか
| 手順 | このツール（WebDriverBiDiCore.cls） | chromium-bidi v17.0.2（MapperCdpConnection.ts） | ChromeDriver |
|---|---|---|---|
| 3. 非表示のタブを開く | [L113-L122](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L113-L122) | [L146-L153](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L146-L153) | [chrome_impl.cc L289-L304](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/chrome_impl.cc;l=289-304) |
| 4. `Target.exposeDevToolsProtocol` | [L124-L129](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L124-L129) | [L180-L184](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L180-L184) | [devtools_client_impl.cc L337-L348](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=337-348) |
| 5. `Runtime.addBinding` | [L211-L223](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L211-L223) | [L186-L195](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L186-L195) | [devtools_client_impl.cc L353-L371](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=353-371) |
| 6. `mapperTab.js` を評価 | [L189-L190](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L189-L190) | [L197-L200](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L197-L200) | [devtools_client_impl.cc L384-L404](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=384-404) |
| 7. `runMapperInstance` | [L192-L193](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L192-L193) | [L203-L206](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L203-L206) | [devtools_client_impl.cc L405-L430](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=405-430) |
| 8. BiDi コマンドを渡す | [L379-L388](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L379-L388)（`Runtime.callFunctionOn`） | [L72-L83](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L72-L83)（`Runtime.evaluate`） | [devtools_client_impl.cc L161-L187](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=161-187)（`Runtime.evaluate`） |
| 9. `Runtime.bindingCalled` を受ける | [L331-L371](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L331-L371) | [L93-L102](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L93-L102)・[L166-L169](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L166-L169) | [devtools_client_impl.cc L89-L107](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=89-107) |

`hidden` と `Target.exposeDevToolsProtocol` は、CDP では実験的（experimental）な機能です（[Target.pdl の定義](https://github.com/ChromeDevTools/devtools-protocol/blob/7245ec551bc9d450170193456ab3965bb7e716cb/pdl/domains/Target.pdl#L103-L175)）。タブの中の `window.cdp`・`window.sendBidiResponse`・`window.onBidiMessage`・`window.runMapperInstance` は、chromium-bidi の [bidiTab.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiTab/bidiTab.ts#L28-L99) で定義されています。

手順8だけは、使う CDP のメソッドが違います。どれもタブの中の `window.onBidiMessage` を呼ぶ点は同じですが、chromium-bidi と ChromeDriver は `onBidiMessage("…")` という式を組み立てて `Runtime.evaluate` で評価します。このツールは、JSON のダブルクォートやエスケープで式が崩れるのを避けるため、引数として渡せる `Runtime.callFunctionOn` を使っています（[CDPContext.cls の jsEval](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/CDPContext.cls#L655-L675)）。
:::

おわかりいただけただろうか...
BiDiを送っているはずなのに何故か内部ではCDPとして送っていることに...

実は、BiDiをCDPに訳しているのは、Chromium のブラウザ本体（C++ のソース）ではありません。
訳しているのは、[TypeScriptベースのプロジェクト](https://github.com/GoogleChromeLabs/chromium-bidi)をビルドした`mapperTab.js`です。このプロジェクト（chromium-bidi）のソースは、今は Chromium のソースツリーの `third_party/chromium-bidi` にあります。Chromium 側の説明にも、「ChromeDriver のための WebDriver BiDi の実装で、mapperTab.js にコンパイルされ、ChromeDriver が Chrome ブラウザへ送り込む」とあります（[README.chromium](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:third_party/chromium-bidi/README.chromium;l=14-16)）。
もう一度言います。Chromium本体のccファイルではなく、jsファイルです。つまり、ビルド済みのブラウザのタブで扱える言語です。

そうです。
主役は「exe」ではなかったのです。本当の主役は`mapperTab.js`。
exeはただ
・BiDi の WebSocket の口を開く
・`mapperTab.js`をタブへ運んでセットアップする
・レスポンスの内容をパススルーする

という、非常にシンプルな役目に降格しました。chromium-bidi の README も、ChromeDriver を「BiDi の WebSocket と HTTP の口を持つ C++ のサーバー」、mapper を「ブラウザの中で動き、BiDi のコマンドを CDP のコマンドに訳すもの」と[分けて説明しています](https://github.com/GoogleChromeLabs/chromium-bidi/blob/c05676b921aa0011c9e124ef91b51eb9dd259e8b/README.md?plain=1#L442-L445)。

今回Excelで再現したこのフローは、[chromium-bidi](https://github.com/GoogleChromeLabs/chromium-bidi)を参考に作り上げました。手本にした Node.js 版の [`src/bidiServer/MapperCdpConnection.ts`](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts) は v17.0.2 まであり、2026年9月に [main から削除されました](https://github.com/GoogleChromeLabs/chromium-bidi/commit/314cf35cfb3de63881fd735b0c3add327d2e2cb2)。
![大まかな手順](/img/story/7.png)

じゃ～WebDriver.exeは、どうなのか...

実は同じフローなのです（違うのは、手順8で `Runtime.evaluate` を使う点くらい）。
なんなら、<準備編>の2に当たるものもexeにあります。巨大な配列定数としての`mapperTab.js`です。役割は、Excelと完全に一致しているのです。
ChromeDriver のビルドでは、chromium-bidi の `mapperTab.js` を縮めてから、`kMapperScript` という C++ の文字配列にして埋め込んでいます（[BUILD.gn](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/BUILD.gn;l=76-104)、[embed_bidimapper_in_cpp.py](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/embed_bidimapper_in_cpp.py;l=33-45)、[cpp_source.py](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/cpp_source.py;l=63-71)）。実行時は、この `kMapperScript` をタブへ送り込みます（[session_commands.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=440-462)）。

![バイナリエディタでの検索結果](/img/story/18.png)

著作権表示と共に、`<!DOCTYPE html><title>BiDi-CDP Mapper</title>...` という生々しいコードが、数万行にわたってハードコードされていた！
ChromeDriver のビルドは、縮めるときも `Copyright` や `license` を含むコメントを残す設定になっています（[minify_js.js](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:ui/webui/resources/tools/minify_js.js;l=70-77)）。`<!DOCTYPE html>...` は、mapper のタブに表示する画面の HTML で、chromium-bidi の [mapperTabPage.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiTab/mapperTabPage.ts#L19-L21) にあります。

exe：配列（C++ の文字配列）として埋め込み
Excel：複数セルとして分割配置
Access：1レコード
Puppeteer：npm の依存として chromium-bidi を入れ、直接呼び出す（ただし、タブに `mapperTab.js` を載せず、[Node.js の中で mapper を動かす](https://github.com/puppeteer/puppeteer/blob/a207faf527ce4305b33a6c0f9706b090c21c115e/packages/puppeteer-core/src/bidi/BidiOverCdp.ts#L59-L66)）

他の有名ツールもWebDriver.exeも自分で組んだVBAも、同じ chromium-bidi の mapper でChromiumをBiDi化していたのだ。
決して無駄なことではなかったと分かり、安心した瞬間だった。
こうして、[v1.6.0](https://github.com/Eschamali/StarterWebScrapingKit/releases/tag/v1.6.0)（2026年3月）で、外部依存なしでWebDriverBiDiが話せるようになりました。
ドライバー（ChromeDriver）を使って VBA から BiDi を扱う試みは、[note の記事](https://note.com/sele_chan/n/n2b21c7c26ef8)や [WebDriver BiDi for SeleniumVBA](https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA) のように先にありました。ドライバーなしで話すものは、私が調べた範囲では見当たらなかったので、VBA 初だと思っています。

![完成](/img/story/12.png)

## Firefoxはどうなのか？

実はこのツール、最近になって、ついにFirefoxの制御にも成功しました。

さっき、WebDriverBiDiも話せるようになった、とお伝えしました。
しかし、あの時はまだかなり密結合な状態で、あくまでもChromium専用のBiDiClassとして作っていました。

何故なら、FirefoxはWebDriverBiDiをネイティブに扱えるからです。Firefox を `--remote-debugging-port` を付けて起動すると、本体に入っている Remote Agent が WebDriver BiDi の口を開きます（[Firefox のソース](https://github.com/mozilla-firefox/firefox/blob/d6fd7a4e69b3c2d4550582eb38d2cdbd9cbad79a/remote/components/RemoteAgent.sys.mjs#L647-L649)、[Firefox のドキュメント](https://firefox-source-docs.mozilla.org/remote/index.html)）。
つまり、このツールの場合は、CDP系のClassを経由せずに直接Firefoxへ届けないと、一切話せないのです。

WebDriverBiDiはクロスブラウザの規格なのに、Chromiumしか話せない...
そんな矛盾が、しばらく続いてました。

しかし、幾度もの改修工事(RaiseEvents、CDPを真似たClass構成、徹底的な責務分離、OptionClass概念...)を経て、ついに...
メインのコードに数十行を追記するだけで、起動済みのFirefoxに限りますが、このツールでもFirefoxを比較的容易に制御できるようになりました。

「Chromium で積み上げた設計が、本物の WebDriver BiDi にもそのまま通用した」という、ちょっとした答え合わせになりました。


## 将来、ChromiumはWebDriverBiDiをネイティブに扱えるのか？

執筆時点では少なくとも、
・Chromiumに、(Firefoxのように)WebDriverBiDiをネイティブに扱える起動オプションのようなもの

は存在しませんし、公式の情報やソースからも、そのような示唆は見当たりませんでした。
Chromium の起動引数を定義している [content_switches.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:content/public/common/content_switches.cc;l=592-599) を見ても、リモートデバッグ用の引数は `--remote-debugging-pipe` や `--remote-debugging-port` などの CDP 用で、BiDi という語は出てきません。`content_features.cc` と `chrome_features.cc` も同じです。
ただ、ChromeDriver のソースには「Mapper をタブの外へ移したら消す」という TODO があり（[chrome_launcher.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome_launcher.cc;l=255-263)）、mapper の置き場所を変える構想自体はあるようです。2026年9月には、chromium-bidi 側も Node.js 版のサーバーを消して、[ChromeDriver 専用になりました](https://github.com/GoogleChromeLabs/chromium-bidi/commit/314cf35cfb3de63881fd735b0c3add327d2e2cb2)。

ここからは私の予想になりますが、
・一生、ネイティブ搭載はしない。`mapperTab.js`経由を強いられる

と思っています。
何故そう思うか？理由はシンプルです。
・代替機能がないから

実はこの辺り、VBAと似た構図があります。

Microsoftはここんところ、いくつかの製品のサポート終了や廃止を進めています。
1. Microsoft Publisher（2026年10月にサポート終了。[告知](https://support.microsoft.com/en-US/publisher/microsoft-publisher-will-no-longer-be-supported-after-october-2026)）
2. VBS（2023年10月に非推奨となり、段階的に削除。[告知](https://learn.microsoft.com/ja-jp/windows/whats-new/deprecated-features)）
3. そしてもちろんIE（2022年6月にサポート終了。[告知](https://learn.microsoft.com/ja-jp/lifecycle/announcements/internet-explorer-11-end-of-support)）

代表に挙げた3つには、どれも代替品があります。
1. Microsoft Create、Word や PowerPoint にあるテンプレートなど
2. ps1（Microsoft は PowerShell と JavaScript への移行を[勧めています](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301)）
3. Edge（IE からの移行先として Microsoft が勧めている、Chromium ベースの Edge）

では、VBAはどうか？
ないんです。

何故か？では、下記の代替機能は他のプログラミング言語にありますか？特にExcelVBAはどうです？
・環境構築不要(Officeのインストーラーを実行すれば、Officeを開いてAlt + F11ですぐ始められる)
　→他言語の場合は、追加でVS Codeが要る

・ワークシートがGUIエディター
　→他言語の場合は、XAMLなど、UIをコードで組み立てる言語の知識が求められる

・データとプログラムが一心同体(ワークシート/Accessテーブル)
　→他言語の場合は、SQLiteやini、csvといった外部ファイルに分けることになり、うっかりファイルを移動・削除なんかしたら、もう動かなくなることも
　　責務分離としては美しいので、ここは賛否両論ありそう。

・ファイルの配りやすさが異常
　→他言語の場合は、ZIPで見た目だけ1ファイルにする必要があり、実行するときは結局ファイルが散らばるし、ランタイム(Python、Node.js)がなかったら詰み
　　まぁ、Officeがなかったらそれも詰みではあるが、「会社」という括りなら相殺される

・追加申請不要(業務遂行上必要なファイルとして情シスの検問をPASS)
　→他言語の場合は申請が要るし、待たされた挙句に拒否されることも。
　　まぁ、これは個人のPCでは意味のない話だが、「会社」という括りなら価値あるメリット

どうです？この5つを全て叶えるプログラミング言語、思いつきます？思いつかないと思います。

意外かもしれませんが、
銀行、証券、工場、官公庁……。
世界中の名だたる大企業の「絶対に止めてはいけない日次処理」が、誰が書いたかも分からないスパゲッティVBAで動いています。
Microsoftがもし「明日からVBAを廃止します」と言った瞬間、経済が止まるほどです。

CDPも実は同じです。思いつきますか？CDPの代替機能。
ないです。だってGoogle（Chromium）自身が育ててきた、プログラミング言語ならぬ、プロトコルなのですから。
Chrome チームも、「WebDriver BiDi は CDP を置き換えない。Chromium 系のブラウザは、デバッグのために CDP を使い続ける」と[書いています](https://developer.chrome.com/blog/webdriver-bidi)。

唯一近いのがWebDriverBiDiですが、これはあくまでも、自動化という括りの中で制定されたものです。
CDPはデバッグ兼自動化。BiDiは自動化だけ。
CDP の[公式ページ](https://chromedevtools.github.io/devtools-protocol/)は、CDP を「デバッグと自動化の対象を、デバッグと自動化のクライアントとつなぐプロトコル」と説明しています。W3C は WebDriver を、「ユーザーエージェントを調べて操作するための[遠隔操作のインターフェース](https://www.w3.org/TR/2018/REC-webdriver1-20180605/)」としています。

どっちがカバー範囲が大きいかは明らかですね。
ただ、それだけのことです。

## `BiDi-CDP Mapper`(mapperTab.js)の未来予想

この技術資料を執筆しているときに、とある[コメント](https://source.chromium.org/chromium/chromium/src/+/main:chrome/test/chromedriver/chrome_launcher.cc;l=259)が目に留まった。
> Remove after the Mapper is moved away from the tab.

Mapper がタブから別の場所へ移動されたら、このコードは削除すること

このコメントは、2024年5月3日の[コミット](https://chromium.googlesource.com/chromium/src/+/f6f5874c8ec33cf232b4fd7aa921fd3ed0e0e7a5)で追加されたものです。

コメントが挙げているのは[4762](https://crbug.com/chromedriver/4762)だけだが、実は[1つ前の4761](https://crbug.com/chromedriver/4761)もこの`BiDi-CDP Mapper`の件を扱っていた。
（このコメントを追加したコミットの説明文では、「Mapper がタブの外で動くようになったら、このスイッチは外せる」という一文に、4762 ではなく 4761 が添えられています）
要約すると
・[4762](https://issues.chromium.org/issues/42323802)「BiDi Mapper tab bring throttled when in background」：バックグラウンドタブだと`BiDi-CDP Mapper`が抑制（スロットリング）され、パフォーマンス低下に
　→専用起動引数`disable-background-timer-throttling`で対処（2024年5月に Fixed）

・[4761](https://issues.chromium.org/issues/42323801)「Hide the BiDiMapper tab in the browser」：`BiDi-CDP Mapper`の非表示方法の模索
　→2025年4月、CDP の`Target.createTarget`に`hidden`が追加され（[crbug 378635448](https://issues.chromium.org/issues/378635448)、[コミット](https://chromium.googlesource.com/chromium/src/+/b6d823369f5ba7439e36b99aa6eb9ef981ae1c3e)）、ChromeDriver も`hidden: true`のターゲットで Mapper を動かすようになった（[コミット](https://chromium.googlesource.com/chromium/src/+/753acf9df872f629667c329dcfaf73677e8d9f93)）
　（4761 自体は 2025年5月に、非公開の別 Issue の重複として閉じられています）

`BiDi-CDP Mapper`に対する格闘が見て取れます。
そして、追加から約2年5か月経った今（2026年10月）でも、このTODOコメントは残ったままです。

ちなみに、2025年11月には同じ`disable-background-timer-throttling`が、BiDi かどうかに関係なく ChromeDriver の全セッション共通の起動引数にも入りました（[コミット](https://chromium.googlesource.com/chromium/src/+/23902d9051098fbc6a3a669fc0988c6899f38b65)、[chromium-bidi#3894](https://github.com/GoogleChromeLabs/chromium-bidi/issues/3894)）。TODO の付いた BiDi 用の分岐は、それとは別に今も残っています。

では、Mapper は一体どこへ移動するのか。考察する。

### 移行先の候補

#### Service Worker
まず移行先の候補だが、ないわけではない。
実は、2021年4月15日付の「[WebDriver BiDi in Chrome Context implementation plan](https://docs.google.com/document/d/1VfQ9tv0wPSnb5TI-MOobjoQ5CXLnJJx9F_PxOMQc8kY)」（実装計画書）に、Mapper を画面から見えなくする方法の1つとして、Service Worker で動かす案が上がっていた。
（この計画書での推奨は「当面は隠さず、見えるタブで動かす」で、Service Worker 案は「要検討」扱い。その前の[下書き版](https://docs.google.com/document/d/1RLIna4f_tZzufieLUA7KTqfp96G46tedKOlWMdRS1BE)では「長期の解として推奨」とされていました）

ユーザーがいじれない領域で稼働するので相性抜群と思いきや、動いていない時には消えてしまう問題があった。
計画書でも短所として「生かし続けるための追加の手間が要るかもしれない」「ワーカーをどう開いたままにするか」が挙がっています。Chrome は何もしていない Service Worker を止めます（例えば拡張機能の Service Worker は、[30秒何もしないと止められます](https://developer.chrome.com/docs/extensions/develop/concepts/service-workers/lifecycle)）。

つまり Service Worker で動かしていたら、例えば VBA で Sleep とかすると`BiDi-CDP Mapper`も消えかねない、ということなのだ...

仮に何とか Sleep させないような仕組みのパッチを当てたとしても、今度は「Target.exposeDevToolsProtocol」のてこ入れが必要になる。
[Target.exposeDevToolsProtocol](https://chromedevtools.github.io/devtools-protocol/tot/Target/#method-exposeDevToolsProtocol)は、対象のメインフレームの`window[bindingName]`に通信用のオブジェクトを差し込む仕組みで、現時点では page ターゲット（タブ。hidden ターゲットも含む）専用です。WebContents を持たないターゲットには「RemoteDebuggingBinding can be granted only to page targets」というエラーで失敗します（[target_handler.cc](https://source.chromium.org/chromium/chromium/src/+/main:content/browser/devtools/protocol/target_handler.cc) の`TargetHandler::ExposeDevToolsProtocol`）。
ここもてこ入れすることになるのだ。

修正量は未知数だが、さきほど挙げた「Target.createTarget - hidden:=True ×専用起動引数で対処」のほうが、修正量としては多分少ないんだろうな～と見た。
（実際、2025年4月以降の ChromeDriver はこの形で動いています）

#### Chromium の C++ ソース（.cc）に`mapperTab.js`の中身を定数として置く

以前お伝えした通り、WebDriver.exe ルートでは、exe の中にある定数をそのままブラウザに投げて`BiDi-CDP Mapper`を起動しています。
（ChromeDriver のビルドでは、[BUILD.gn](https://source.chromium.org/chromium/chromium/src/+/main:chrome/test/chromedriver/BUILD.gn) の`minify_bidimapper`と`embed_bidimapper_in_cpp`で、`mapperTab.js`を縮めてから C++ の定数として埋め込んでいます）

じゃ～その定数をWebDriver.exeじゃなくて、Chrome本体（Chromiumソース）に丸ごと移せばいいじゃん？と思いますよね。でもダメなんです。

実はこの案、2022年1月の設計文書「[CDP-BiDi Mapper code lifecycle](https://docs.google.com/document/d/1_UTBlOa5GCdf075ni4hDVgBSHkgeblfP_g_hwxqRhjg)」で「Hardcode in Chrome」（Mapper のソース文字列を Chrome 自体と一緒にコンパイルし、CDP コマンドで起動する案）として検討済みです。そこで挙がった難点は次の3つです。

・バイナリサイズ：「約30KB増え、それが妨げになりうる」
・承認の難しさ：「そのような CDP コマンドを正当化して承認を得るのは危うく、できたとしても長くかかりうる」
・依存の向き：「Mapper の考えを CDP に持ち込むと依存の向きが逆になり、CDP が Mapper を知ることになる」

結果、2022年の推奨は「Mapper は ChromeDriver・Puppeteer 側に定数として持たせる」でした。

バイナリサイズの話でいうと、これはうちも初めて気づいたのですが、Chromium には[バイナリサイズを見張る仕組み](https://chromium.googlesource.com/chromium/src/+/main/docs/speed/binary_size/binary_size_explainer.md)があります。例えば[android-binary-size トライボット](https://chromium.googlesource.com/chromium/src/+/main/docs/speed/binary_size/android_binary_size_trybot.md)は、1コミットで Android の APK が 16KB（arm32）を超えて増えるとチェックに引っかかります（減らせない場合は、増える理由をコミットに書いて通す運用です）。

そして、設計文書の見積もりは約30KBでしたが、今の`mapperTab.js`は約1MBあります（このリポジトリ同梱版で 868,468 バイト、npm の最新版 157.0.8103-20 で 1,016,727 バイト）。当然、埋め込めばその分ブラウザのサイズも膨れ上がります。
しかも`mapperTab.js`は、自動化（テスト）でしか使わないコードです。一般の利用者が使わないコードのために、全員が使うブラウザを大きくするのは筋が悪い、というのは想像しやすいと思います。

あとは、Mapper が CDP と W3C-WebDriverBiDi の両方の破壊的変更の影響を受けること（設計文書の「Version pinning」の節）も、この定数配置の壁になっています。

なお2026年8月、chromium-bidi のソース自体は Chromium のリポジトリ（[third_party/chromium-bidi](https://chromium.googlesource.com/chromium/src/+/main/third_party/chromium-bidi)）に引っ越し、ChromeDriver はそのソースから`mapperTab.js`をビルドして埋め込むようになりました（[コミット](https://chromium.googlesource.com/chromium/src/+/034e431ee2b07ebc99818691019052a161bea2ed)）。ソースの置き場所は Chromium になりましたが、Mapper をブラウザに送り込むのは今も ChromeDriver です。「Chromium のソースに移す」と「ブラウザ本体に埋め込む」は別物、というわけです。

### Firefoxの場合は？

Firefoxは、WebDriver BiDi をブラウザ本体の中に実装している。その BiDi のモジュールは`.sys.mjs`、つまり JavaScript で書かれていて（[MessageHandler](https://firefox-source-docs.mozilla.org/remote/messagehandler/index.html)という枠組みの上で動く）、コンパイルの要る部分は C++ と Rust（[Building the Remote Agent](https://firefox-source-docs.mozilla.org/remote/Building.html)）。高速な C++/Rust と、拡張しやすい JS のハイブリッド運用といったところだ。しかも`.sys.mjs`はビルドし直さずに変更できる。
一方Chromium系は、高速を売りにしているのでコアはほぼ C++。`mapperTab.js`を定数として置こうとすると、おのずと C++ の文字列定数一択になる（実際 ChromeDriver でも C++ の定数にしている）。するとどうなると思う？

W3C-WebDriverBiDiに更新→定数置き換え→そのたびにビルド時間という待ち時間発生→タイムパフォーマンス低下
結果：`mapperTab.js`は、WebDriver.exe専用として埋め込み。WebDriver.exe自体はそこまでのソース量ではないので、ビルド時間もかからない→タイパ良い　ってこと
（※このビルド時間の話は筆者の推測です。設計文書が挙げている理由は、上の節のバイナリサイズ・承認の難しさ・依存の向きです）

### 結論

Chromium 側は、`BiDi-CDP Mapper`の稼働場所をタブから移動しようとは考えている。
でも現時点では、その最適な場所が存在しない。

これが執筆時点での結論と考えています。

とは言え、「稼働場所をタブから移動しよう」という思想がある以上、いつかタブではないどこかで`BiDi-CDP Mapper`が動くことには変わりありません。
2022年の設計文書でも、長期案として「Chrome Component」が「長期の解の良い候補になりうるが、もっと調べる必要がある」とされ、承認者の1人は「Chrome に同梱・固定するほうが長期の選択肢としては魅力的」とコメントしています。

明日なのか数か月、数年後なのか、はたまた一生タブに居続けるのか...

真相は、Google社員のみぞ知る...


## おわりに

ただ、1つだけ言えるのは、Chromium系をBiDiで動かすのに必要なのが、exeというバイナリファイルではなく、JavaScriptというテキストファイルだったことです。

Chrome for Developers のブログ記事「[WebDriver BiDi - クロスブラウザ自動化の未来](https://developer.chrome.com/blog/webdriver-bidi?hl=ja)」（Jecelyn Yeen・Maksim Sadym）に、次の説明と図があります（強調は筆者）。

> WebDriver BiDi のビジョンは、**お気に入りのツール**を使用してテストを作成したり、任意のブラウザまたはドライバでテストを自動化したりすることで、柔軟性を最大限に高めることです。

このビジョンの一文には、Puppeteer といったGoogle製ツールの名前は出てきません。汎用的に制御できる仕組みを提供しなければ、という意思があったのだと思います。

そしてその**汎用**を担うのが、どのPC/スマートフォン/OSにも必ず入っているブラウザと、その中で動かせるスクリプト言語である JavaScript です。
（ブラウザで動く言語には、2019年に W3C 勧告となった[WebAssembly](https://www.w3.org/press-releases/2019/wasm/)もありますが、ここで言いたいのはテキストのまま扱えるスクリプトとしての JavaScript です）
WebDriver.exeといったコンパイル済み実行形式バイナリだと、OS依存になってしまいます。

では、なぜ Chromium系をBiDiで動かすコアは JavaScript になったのか。2021年の実装計画書が挙げている理由は「Mapper を C++ で実装せずに済み、JavaScript の力で非同期のコードを書ける」ことでした。ほかの案は、次の理由で退けられています。

・ChromeDriver に実装する案：「ChromeDriver の大幅な改修が要る」
・Chromium に実装する案：「Chromium のバイナリが大きく増えうる」「ブラウザプロセスで信頼できないデータを C++ で扱う安全上の懸念」
・Rust で実装する案：「Chromium の基盤がまだ Rust に対応していない」

OS に依存しないから、という理由は計画書には書かれていません。それでも結果として、OS を問わず同じ JavaScript を使い回せる形になったのだと思います。

JavaScriptは、ただのテキストファイルです。だから、Excelのセルに埋め込めました。Accessなら、レコードに入れられました。
JavaScriptは基本、ブラウザ専用言語です。だから、exeみたいに一旦デスクトップに展開する必要がなく、埋め込んだJavaScriptテキストをブラウザに直接入れるだけで、BiDiを扱えました。

このフローは別に、VBAに限りません。
違うのは、`mapperTab.js`の置き場所だけです。
Office系ならそのファイル内に入れられるし、CUI系言語（Node.js、Python等）なら単純にファイルにしてデスクトップに置けば済みます。
実際、Puppeteer は Mapper（chromium-bidi）を npm パッケージとして取り込み、Node.js 側で動かす仕組みも持っています（[BidiOverCdp.ts](https://github.com/puppeteer/puppeteer/blob/main/packages/puppeteer-core/src/bidi/BidiOverCdp.ts)）。

そして、
・クロスプラットフォーム対応プログラミング言語×クロスプラットフォーム対応の`mapperTab.js`
→`mapperTab.js`がWebDriver.exe依存だったら、実質クロスプラットフォームになれない

こうなっていたからこそ、より**お気に入りのツール**を使用して自動化しやすくなったのではないでしょうか？
そしてその**お気に入りのツール**には、奇跡的にVBAも含まれていました。

また、「任意のブラウザまたはドライバ」と書いてあるあたり、WebDriver.exe依存じみた書き方にはなっていません。
mapperTab.jsの注入/起動方法さえ分かれば、わざわざWebDriver.exeを介さなくても、どのプログラミング言語でもChromium系に直接BiDiを送れます。VBAも例外ではないのです。

CDP⇔BiDi翻訳という仕組みを、最初は「なんて泥臭いその場凌ぎなんだ」と思っていましたが、今となっては、mapperTab.**exe**(WebDriver.exe)ではなく、mapperTab.**js**として登場してよかったと思っています。
（ちなみに、Puppeteer で BiDi サーバーを作る案の[設計文書](https://docs.google.com/document/d/1FAaZ4RS1fatr5aPJJNW-PkZREHv5tfJOb6ao_j9dI2Y)では、JS の Mapper は中期の解で、長期的には C++ や Rust などに置き換える想定でした。当の Google も、最初はその場凌ぎのつもりだったのかもしれません）

`BiDi-CDP Mapper`、タブの次はどこに向かうのか、一緒に行方を見守りましょう。

## 参考資料

ソースコードへのリンクは、確かめた時点（2026年10月）の版に固定しています。

### Chromium・ChromeDriver・chromium-bidi

* [third\_party/chromium-bidi/README.chromium](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:third_party/chromium-bidi/README.chromium;l=1-16)（Chromium の中の chromium-bidi。Version 17.0.2。「mapperTab.js にコンパイルされ、ChromeDriver が Chrome へ送り込む」）

* [chrome/test/chromedriver/BUILD.gn](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/BUILD.gn;l=76-104)（`mapperTab.js` を縮めて C++ に埋め込む）

* [embed\_bidimapper\_in\_cpp.py](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/embed_bidimapper_in_cpp.py;l=33-45)・[cpp\_source.py](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/cpp_source.py;l=63-71)（`kMapperScript` という文字配列にする）

* [ui/webui/resources/tools/minify\_js.js](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:ui/webui/resources/tools/minify_js.js;l=70-77)（縮めるときに著作権のコメントを残す設定）

* [chrome/test/chromedriver/session\_commands.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=411-475)（非表示タブ、`kMapperScript`、`session.new`）

* [chrome/test/chromedriver/chrome/devtools\_client\_impl.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome/devtools_client_impl.cc;l=303-430)（`StartBidiServer`。mapper を起動する手順）

* [chrome/test/chromedriver/chrome\_launcher.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome_launcher.cc;l=255-263)（BiDi は専用タブの mapper で動く。タブの外へ移す TODO）

* [content/public/common/content\_switches.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:content/public/common/content_switches.cc;l=592-599)（リモートデバッグの起動引数）

* [chromium-bidi v17.0.2 の MapperCdpConnection.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts#L137-L210)（Node.js から mapper のタブを立ち上げる手順。このツールの手順の元）

* [chromium-bidi v17.0.2 の bidiTab.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiTab/bidiTab.ts#L28-L99)（タブ側の `runMapperInstance`・`onBidiMessage`）

* [chromium-bidi v17.0.2 の mapperTabPage.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiTab/mapperTabPage.ts#L19-L21)（`<!DOCTYPE html><title>BiDi-CDP Mapper</title>`）

* [chromium-bidi の README](https://github.com/GoogleChromeLabs/chromium-bidi/blob/c05676b921aa0011c9e124ef91b51eb9dd259e8b/README.md?plain=1#L1-L10)（BiDi と CDP の間を訳す JavaScript の層で、Chrome のタブの中で動く。ソースは Chromium に移った）

* [\[chromium-bidi\] Remove Node.js BiDi server and use ChromeDriver only](https://github.com/GoogleChromeLabs/chromium-bidi/commit/314cf35cfb3de63881fd735b0c3add327d2e2cb2)（2026年9月）

* [Target.pdl](https://github.com/ChromeDevTools/devtools-protocol/blob/7245ec551bc9d450170193456ab3965bb7e716cb/pdl/domains/Target.pdl#L103-L175)（`Target.exposeDevToolsProtocol` と `Target.createTarget` の `hidden`）

* [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/)

* [Puppeteer の BidiOverCdp.ts](https://github.com/puppeteer/puppeteer/blob/a207faf527ce4305b33a6c0f9706b090c21c115e/packages/puppeteer-core/src/bidi/BidiOverCdp.ts#L59-L66)（Node.js の中で chromium-bidi の BidiServer を起動する）

### 仕様

* [WebDriver BiDi（W3C Editor's Draft）](https://w3c.github.io/webdriver-bidi/)（[Establishing a Connection](https://w3c.github.io/webdriver-bidi/#establishing)、[session.new](https://w3c.github.io/webdriver-bidi/#command-session-new)）

* [WebDriver BiDi の explainer](https://github.com/w3c/webdriver-bidi/blob/f922e6af669ead51496db85f37813f39438ef171/explainer.md)（TPAC 2019 で話し合った設計案）

* [WebDriver（W3C 勧告、2018年6月5日）](https://www.w3.org/TR/2018/REC-webdriver1-20180605/)

### 歴史

* [Selenium History](https://www.selenium.dev/history/)

* [Selenium RC（Selenium 1）の公式ドキュメント](https://www.selenium.dev/documentation/legacy/selenium_1/)（Proxy Injection）

* [Selenium WebDriver（The Architecture of Open Source Applications 第16章、Simon Stewart）](https://aosabook.org/en/v1/selenium.html)

* [Introducing WebDriver（Google Open Source Blog、2009年5月、Simon Stewart）](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)

* [Selenium 2.0: Out Now!（2011年7月、Simon Stewart）](https://www.selenium.dev/blog/2011/selenium-2-0/)

* [Selenium 2.0 alpha 1 の Chrome ドライバーの manifest-win.json](https://github.com/SeleniumHQ/selenium/blob/05ceaf552081cba0035862ea0483c08b2692338e/chrome/src/extension/manifest-win.json#L4-L17)（拡張機能と NPAPI のプラグイン）

* [WebKit Remote Debugging（2011年5月、Pavel Feldman。Wayback Machine）](https://web.archive.org/web/20110513094715/http://www.webkit.org:80/blog/1620/webkit-remote-debugging/)

* [Announcing Remote Debugging Protocol v1.0（2012年4月、Pavel Feldman。Wayback Machine）](https://web.archive.org/web/20120507025807/http://www.webkit.org/blog/1875/announcing-remote-debugging-protocol-v1-0/)

* [A look back in time: the evolution of test automation（Chrome for Developers）](https://developer.chrome.com/blog/test-automation-evolution)

* [WebDriver BiDi - The future of cross-browser automation（Chrome for Developers）](https://developer.chrome.com/blog/webdriver-bidi)

* [WebDriver BiDi production-ready in Firefox, Chrome and Puppeteer（Chrome for Developers、2024年8月）](https://developer.chrome.com/blog/firefox-support-in-puppeteer-with-webdriver-bidi)

* [Removing ChromeDevTools Support For Firefox（Selenium、2025年）](https://www.selenium.dev/blog/2025/remove-cdp-firefox/)

* [Headless Chrome（Chrome for Developers）](https://developer.chrome.com/blog/headless-chrome/)

* [YouTube now defaults to HTML5 \<video\>（2015年1月。Wayback Machine）](https://web.archive.org/web/20160101205917/http://youtube-eng.blogspot.com/2015/01/youtube-now-defaults-to-html5_27.html)

### Firefox・Microsoft

* [Firefox の RemoteAgent.sys.mjs](https://github.com/mozilla-firefox/firefox/blob/d6fd7a4e69b3c2d4550582eb38d2cdbd9cbad79a/remote/components/RemoteAgent.sys.mjs#L647-L649)（`--remote-debugging-port` で WebDriver BiDi の Remote Agent を起動する）

* [Remote Protocols（Firefox Source Docs）](https://firefox-source-docs.mozilla.org/remote/index.html)

* [Microsoft Edge DevTools Protocol](https://learn.microsoft.com/ja-jp/microsoft-edge/devtools/protocol)

* [Excel ワークシートとブックの制限](https://support.microsoft.com/ja-jp/office/excel-worksheet-and-workbook-limits-c13798fc-f1d5-4e73-998d-c02268bc65ae)（セルの文字数 32,767）

* [Access デスクトップ データベースのデータ型](https://support.microsoft.com/ja-jp/access/data-types-for-access-desktop-databases)（長いテキストは最大 1 GB）

* [Microsoft Publisher will no longer be supported after October 2026](https://support.microsoft.com/en-US/publisher/microsoft-publisher-will-no-longer-be-supported-after-october-2026)

* [Windows クライアントの非推奨の機能](https://learn.microsoft.com/ja-jp/windows/whats-new/deprecated-features)（VBScript、2023年10月）

* [VBScript deprecation: Timelines and next steps（Windows IT Pro Blog、2024年5月）](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301)

* [Internet Explorer 11 のサポート終了](https://learn.microsoft.com/ja-jp/lifecycle/announcements/internet-explorer-11-end-of-support)

### VBA から BiDi へ

* [GCuser99/SeleniumVBA Discussions #164](https://github.com/GCuser99/SeleniumVBA/discussions/164#discussioncomment-15988897)（2026年3月）

* [SeleniumVBA](https://github.com/GCuser99/SeleniumVBA)

* [\[SeleniumVBA\]WebDriver BiDiの接続にチャレンジしてみた（note、2025年6月）](https://note.com/sele_chan/n/n2b21c7c26ef8)

* [WebDriver BiDi for SeleniumVBA](https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA)

* [このツールの WebDriverBiDiCore.cls](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/WebDriverBiDiCore.cls#L109-L134)（`newTabForBiDi`・`StartBiDiCDPMapper`・`RunCDP`）

* [v1.6.0 のリリース](https://github.com/Eschamali/StarterWebScrapingKit/releases/tag/v1.6.0)（2026年3月15日）

- Chromium: [chrome_launcher.cc](https://source.chromium.org/chromium/chromium/src/+/main:chrome/test/chromedriver/chrome_launcher.cc)（TODO コメント、2026年10月時点の main）
- Chromium コミット: [f6f5874](https://chromium.googlesource.com/chromium/src/+/f6f5874c8ec33cf232b4fd7aa921fd3ed0e0e7a5)「Do not throttle background tabs in BiDi mode」（2024-05-03）
- Chromium コミット: [b6d8233](https://chromium.googlesource.com/chromium/src/+/b6d823369f5ba7439e36b99aa6eb9ef981ae1c3e)「Create hidden targets via CDP `Target.CreateTarget`」（2025-04-02）
- Chromium コミット: [753acf9](https://chromium.googlesource.com/chromium/src/+/753acf9df872f629667c329dcfaf73677e8d9f93)「Reland "ChromeDriver runs BiDi Mapper in a hidden target"」（2025-04-09）
- Chromium コミット: [23902d9](https://chromium.googlesource.com/chromium/src/+/23902d9051098fbc6a3a669fc0988c6899f38b65)「Disable backgrounding occluded windows for ChromeDriver sessions」（2025-11-06）
- Chromium コミット: [034e431](https://chromium.googlesource.com/chromium/src/+/034e431ee2b07ebc99818691019052a161bea2ed)「Integrate chromium-bidi into chromedriver」（2026-08-24）
- Issue: [chromedriver/4761](https://issues.chromium.org/issues/42323801)、[chromedriver/4762](https://issues.chromium.org/issues/42323802)、[378635448](https://issues.chromium.org/issues/378635448)
- 設計文書: [WebDriver BiDi in Chrome Context implementation plan](https://docs.google.com/document/d/1VfQ9tv0wPSnb5TI-MOobjoQ5CXLnJJx9F_PxOMQc8kY)（Maksim Sadym、2021-04-15）
- 設計文書: [Draft: WebDriver BiDi in Chrome Page/Extension implementation plan](https://docs.google.com/document/d/1RLIna4f_tZzufieLUA7KTqfp96G46tedKOlWMdRS1BE)
- 設計文書: [CDP-BiDi Mapper code lifecycle](https://docs.google.com/document/d/1_UTBlOa5GCdf075ni4hDVgBSHkgeblfP_g_hwxqRhjg)（Maksim Sadym、2022-01-03）
- 設計文書: [WebDriver BiDi Server using Puppeteer](https://docs.google.com/document/d/1FAaZ4RS1fatr5aPJJNW-PkZREHv5tfJOb6ao_j9dI2Y)
- 設計文書の一覧: [ChromeDriver documentation](https://developer.chrome.com/docs/chromedriver/documentation)
- CDP: [Target ドメイン](https://chromedevtools.github.io/devtools-protocol/tot/Target/)（`createTarget`の`hidden`、`exposeDevToolsProtocol`）
- Chromium: [Binary Size Explainer](https://chromium.googlesource.com/chromium/src/+/main/docs/speed/binary_size/binary_size_explainer.md)、[Trybot: android-binary-size](https://chromium.googlesource.com/chromium/src/+/main/docs/speed/binary_size/android_binary_size_trybot.md)
- Chrome for Developers: [The extension service worker lifecycle](https://developer.chrome.com/docs/extensions/develop/concepts/service-workers/lifecycle)
- Chrome for Developers: [WebDriver BiDi - クロスブラウザ自動化の未来](https://developer.chrome.com/blog/webdriver-bidi?hl=ja)（Jecelyn Yeen・Maksim Sadym）
- Firefox Source Docs: [Building the Remote Agent](https://firefox-source-docs.mozilla.org/remote/Building.html)、[MessageHandler](https://firefox-source-docs.mozilla.org/remote/messagehandler/index.html)
- Puppeteer: [BidiOverCdp.ts](https://github.com/puppeteer/puppeteer/blob/main/packages/puppeteer-core/src/bidi/BidiOverCdp.ts)
- W3C: [WebAssembly becomes a W3C Recommendation](https://www.w3.org/press-releases/2019/wasm/)（2019-12-05）

## 次に読む

- [アーキテクチャ](/concepts/architecture) — mapper が載る層の位置づけ
- [CDP と BiDi](/concepts/cdp-vs-bidi) — どちらを使うか
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol) — `ExecuteBiDi` / BiDi+
