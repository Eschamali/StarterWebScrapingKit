---
description: WebDriver なしでWebDriverBiDiが扱えるとわかった経緯。公式ドライバーに頼らない BiDi 対応の登場秘話です。
---

# なぜ外部依存なしで、ChromiumにWebDriverBiDiが扱えるとわかったのか


![海外フォーラムでの議論](/img/story/1.png)
すべては[この議論](https://github.com/GCuser99/SeleniumVBA/discussions/164#discussioncomment-15988897)から始まりました。

> Well, I think we had this discussion earlier regarding BiDi. CDP is not futureproof as BiDi is.
> And when it comes to companies not allowing local software installations: if Chrome is not being allowed to install, then CDP is worthless anyway. So even in a scenario of restricted environments I assume CDP support not to be a breakthrough.
さて、この話は以前にもBiDiについて話したと思います。CDPはBiDiほど将来性に欠けています。
そして、企業がローカルソフトウェアのインストールを許可しない場合、Chromeのインストールが許可されていなければ、CDPは結局無意味です。ですので、制限された環境のシナリオでも、CDPサポートは画期的なものではないと考えます。

この頃になると、CDPに対する知識は随分と溜まってきました。
なので、すぐに気付きました。半分あっていて、半分間違ってる

CDPは、Chromeブランドではなく、Chromiumブランドです。CDP の[公式ページ](https://chromedevtools.github.io/devtools-protocol/)も、対象を「Chromium、Chrome、そのほかの Blink 系ブラウザ」と書いています。よって、業務PCにプリインストールされているEdgeもCDP対応するということになります。Microsoft も、Edge の DevTools Protocol は Chrome DevTools Protocol の API と一致すると[説明しています](https://learn.microsoft.com/ja-jp/microsoft-edge/devtools/protocol)。
今思えば、あのMicrosoftのEdgeごり押し戦法

・世間からMicrosoftに批判ラッシュがあっても
・大手ニュースから批判されても
・アンインストールできる仕組みはあるのに、アンインストールさせない罠によるバッシング

それでもMicrosoftはWindowsのEdgeを推し進めた
結果的にブック1枚でブラウザ自動化が、IEの頃から(偶然にも？)受け継いだのです。
これは紛れもない事実です。

でも...**「将来性はBiDiが上」**
これも、ある意味、事実です。

その言葉が頭から離れなかったのです。

こうして、私は決めたのです。
>絶対に、外部依存なしで、ChromiumにWebDriverBiDiで扱える方法を探してやる


と
本題に入る前に少しこのWebDriverBiDiに至るまでの歴史をサクッと述べようと思います。

## WebDriverの歴史

立ち位置としてはCDPと同じです。指定のフォーマットに従ってJson文字列を書いてブラウザに送る...違いは誰がルールを決めてるかです。
CDP:Google（Chromium）
WebDriverBiDi:W3C

W3Cは簡単に言うと、大手のブラウザメーカーたちが集まって共通のコマンド名でうちらのブラウザを操る世界標準を作るグループみたいなもんです。

全ての始まりは、GoogleでもなくMicrosoftでもなく、Seleniumです。2004年に ThoughtWorks の Jason Huggins が作ったテストツールが始まりでした（[Selenium 公式の歴史](https://www.selenium.dev/history/)）。
そうです。VBAでは標準となってる`SeleniumVBA`もSeleniumというプロジェクトの上で成り立っています（Selenium の WebDriver を使い、ドライバーが無ければ自動で落としてきます。[SeleniumVBA の README](https://github.com/GCuser99/SeleniumVBA)）

最初の自動化はどうしてたのかというと...
テスト用の JavaScript（Selenium Core）を、テストしたいサイトと同じサーバーに置いて、ブラウザの中で動かしていました。ブラウザには、JavaScript からはページを配ったサーバーにしかアクセスさせない決まり（同一オリジンポリシー）があるからです（[Selenium WebDriver の作者 Simon Stewart による解説](https://aosabook.org/en/v1/selenium.html)）。

でも、テストのたびに相手のサーバーへ置くのは現実的ではありません。そこで登場した Selenium RC では...
なんと！ネットワークの書き換えです🫨

要するに、間にプロキシサーバーを挟んでその内部で受信したhtmlにテストツールのjsを仕込む（[Selenium RC の公式ドキュメント](https://www.selenium.dev/documentation/legacy/selenium_1/)でいう「Proxy Injection」）

という今見たらもはやウイルスチックなことをしてました。でも当時はこれしか選択肢がなかったのです🙃
しかし、当然この仕組みは限界があります。
あくまでも、主体はJavaScriptなので、JavaScriptのルール違反になるようなことをしようとすると当然動きません（ファイルのアップロードや、ドメインをまたぐ移動が例に[挙がっています](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)）。
でも、そのルール違反的なことをしないと、テストもままにならないのも事実...

そして時が経ちようやくここでおなじみのWebDriverが登場しました...いや厳密には「WebDriver」という製品名が来ました（2007年の初めに最初のコードが公開。[Simon Stewart による解説](https://aosabook.org/en/v1/selenium.html)）

でも内部ロジックとしては今と同じように
・各ブラウザごとの操作内容を隠ぺい/共通化
しています。

例えばこのツールのIntroとして出してるIEならCOM
別ブランチ(ex-Firefox)で密かに対応してるFirefoxなら、当時は拡張機能（XPCOM）
今主流のChromeなら、当時は拡張機能＋NPAPI プラグイン

IE と Firefox のやり方は、作者の Simon Stewart 自身が[ブログ](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)と[解説](https://aosabook.org/en/v1/selenium.html)に書いています（IE は MSHTML と ShDocVw の COM インターフェース、Firefox は XPCOM を使う拡張機能）。Chrome は、Selenium 2.0 alpha 1（2009年）の Chrome ドライバーが「WebDriver extension for Chrome」という拡張機能で、[NPAPI のプラグイン `npchromedriver.dll` を積んでいます](https://github.com/SeleniumHQ/selenium/blob/05ceaf552081cba0035862ea0483c08b2692338e/chrome/src/extension/manifest-win.json#L4-L17)。

当然こんなバラバラ相手に1つ1つ学習してたらキリがありません。
それをWebDriverという製品が解決してくれました

そしてSeleniumは気付きました
・うちのJavaScriptハックよりもこのWebDriverのほうがええやん🙃

こうしていい感じに統合されて（2009年に[統合を発表](https://opensource.googleblog.com/2009/05/introducing-webdriver.html)、2011年に [Selenium 2.0 として公開](https://www.selenium.dev/blog/2011/selenium-2-0/)）、2018年、ここでようやく、規格としてのWebDriverが[W3C 勧告](https://www.w3.org/TR/2018/REC-webdriver1-20180605/)として登場しました。
複数ブラウザ相手なのに覚えるAPIはWebDriverという巨大なでも1種類だけ
当時は画期的でした。

しかし、Webの進化は異常に進みました。
ちょっと思い出してみてください。

昔、ブラウザでゲームと言えば
・Flash Player
・Java

みたいな今でいう拡張機能を入れる必要がありました。
「これを使うには、Flash Playerをインストール...」
一体何回見たことやら...

また、当時のYouTubeも実は動画部分をFlash Playerに依存してました（YouTube が HTML5 のプレーヤーを既定にしたのは[2015年1月](https://web.archive.org/web/20160101205917/http://youtube-eng.blogspot.com/2015/01/youtube-now-defaults-to-html5_27.html)）

でも今のブラウザはどうです？
・URLを共有するだけで動画もゲームもすぐにできる/見れる

理由は単純です。それらをつかさどる機能がブラウザに丸ごと次々と標準搭載されたからです
・WebGL API
・Web Audio API
・WebAssembly
・Canvas
...AndMore

こういった技術の結集が高速に実装され、リンク共有だけで素晴らしい体験ができました。

...しかし、この進化に追いつけてないのが1つだけありました。
はい、WebDriverです。

WebDriverの基本は、送って応答待って回答を得るという同期スタイルです。(技術的にいうとhttpといった同期やり取り)
Chrome チームも、WebDriver "Classic" は HTTP が土台なので、ブラウザのイベントを購読して受け取るのが難しいと[書いています](https://developer.chrome.com/blog/test-automation-evolution)。
しかしさきほどの進化の影響でこの仕組みに限界が出てきました
・ページは読み込み終わったのに、ページに表示されてる内容が書き換わる
　→検索ボタンを押したら、まるごとページ読み見直しではなく、部分的に読み直す

・inputタグの動的生成対応
　→初期状態にはなく、ボタン押下することで初めて出現するタイプ。検索では見つからない

・ネットワークの監視
　→非同期処理によるデータ内容に応じた処理

いずれもこれは、SPA（SinglePageApplication）サイトによる「内部で起こった出来事」挙動です。技術が発展したからこそ登場したサイトです。

しかし実はこれら3つとも解決できる手段はあったっちゃありました。
はい、このツールの主題であるCDPです。コンソールのメッセージやネットワークのリクエストの横取りといった低いレベルの操作ができ、WebSocket で双方向にやり取りできます（[同じ解説](https://developer.chrome.com/blog/test-automation-evolution)）。

元々はデバッグ用のプロトコルです。今の CDP につながるのは、2011年に WebKit の Web Inspector（開発者ツール）をブラウザの外からつなぐために[発表](https://web.archive.org/web/20110513094715/http://www.webkit.org:80/blog/1620/webkit-remote-debugging/)された「WebKit Remote Debugging Protocol」です（`Chromium --remote-debugging-port=9222` で起動する例付き。v1.0 は[2012年](https://web.archive.org/web/20120507025807/http://www.webkit.org/blog/1875/announcing-remote-debugging-protocol-v1-0/)）。意外にも、WebDriver（2007年）よりも後発です😳
いまの CDP の定義は [Chromium のソースの中](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:third_party/blink/public/devtools_protocol/)にあります。当然、Webの進化とともに、CDPも同時に進化していきました。
そして進化していくうえで気づいてしまいました
>あれ？これブラウザ自動化にも使えるじゃん

これが、あのPuppeteerの登場です。Chrome チームが作った Node のライブラリで（[Headless Chrome の解説](https://developer.chrome.com/blog/headless-chrome/)）、CDP は「もとは DevTools とデバッグのために作られ、Puppeteer が自動化に取り入れた」と Chrome チーム自身が[書いています](https://developer.chrome.com/blog/test-automation-evolution)。

しかし、名称の通り、あくまでもChromium(Chrome/Edge)向けのプロトコルです。Firefox も一時期、Puppeteer のために CDP の一部を実装していましたが、恒久策ではないとして WebDriver BiDi に移っています（[Chrome チームの記事](https://developer.chrome.com/blog/firefox-support-in-puppeteer-with-webdriver-bidi)。Selenium も 4.29.0 で [Firefox 向けの CDP 対応を外しました](https://www.selenium.dev/blog/2025/remove-cdp-firefox/)）。
このままでは、またあの
・各ブラウザ相手1つ1つ学習してたら鬱になる

が戻ってしまいます

こうして、現在のWebDriver BiDiの制定にいたりました。2019年の W3C の会議（TPAC）で[設計が話し合われ](https://github.com/w3c/webdriver-bidi/blob/f922e6af669ead51496db85f37813f39438ef171/explainer.md)、[仕様](https://w3c.github.io/webdriver-bidi/)は「WebDriver に双方向の通信を足し、ブラウザからイベントを流せるようにする」ものとして書かれています。

## ChromiumをWebDriverBiDiとして話す方法

さて、お待たせしました。ようやく、本題です。
大まかな歴史をつかんだところで、このツールはどうやって、ChromiumをBiDiとして喋っているのか...

### WebDriver.exeルート

とその前に、一般的にWebDriver.exe経由でChromiumに、このWebDriverBiDiを扱うにはどうするのか？
ざっくりフローとしてはこうです
1. DLしたWebDriverを、「--port=9222」のようにポートを指定して起動する
2. そのポートの`/session`に、capabilitiesに「webSocketUrl: true」を付けて送る（New Session）
3. すると ChromeDriver は内部で、mapper（後述）用の非表示タブを立ち上げ、そこへ`session.new`を送ってWebDriverBiDiを有効にする。この`session.NewResult`は ChromeDriver の中で受け取る
4. New Session のレスポンスの capabilities にある`webSocketUrl`に、`ws://<ホスト>/session/<セッションID>`が来るので、そこにWebSocket接続を行う
5. あとはBiDiのルールとしてコマンドのやり取りを、WebSocket上で行う

2〜4は、[W3C の仕様](https://w3c.github.io/webdriver-bidi/#establishing)（WebSocket URL の capability を true で頼むと双方向の接続になる）と、ChromeDriver のソース（[非表示タブを作る所](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=411-418)、[`session.new` を送る所](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=464-475)、[`webSocketUrl` を返す所](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=288-291)）で確かめられます。なお仕様の`session.new`は、HTTP を使わずに BiDi だけでセッションを作るためのコマンドです（[仕様](https://w3c.github.io/webdriver-bidi/#command-session-new)）。

WebDriver.exeで、BiDiを扱う詳細なフローは[ここで](https://note.com/sele_chan/n/n2b21c7c26ef8)確認できます。

### このツール

では、このツールの場合はどうなのか...
大きく2フェーズあります。
こうなります。

<準備編>
1. npm等で`mapperTab.js`を入手（[chromium-bidi](https://www.npmjs.com/package/chromium-bidi) のパッケージに入っています）
2. 中身をExcel/Accessに保存
　→Excelの場合は、セル内に保持できる総文字数が、[32,767 文字](https://support.microsoft.com/ja-jp/office/excel-worksheet-and-workbook-limits-c13798fc-f1d5-4e73-998d-c02268bc65ae)なので、百万字近くある`mapperTab.js`（v17.0.2 で約85万字）を、その文字数ずつ分割保存してます。
　　Accessは、1レコードに収める「長いテキスト」型（[最大 1 GB](https://support.microsoft.com/ja-jp/access/data-types-for-access-desktop-databases)）があるので、特に難なく入れれた

<実際に動かす>
1. BiDiとして使用を開始する`Sub`を呼び、Chromiumのデバッグ(Port/Pipe)に接続
2. 準備編で保存した`mapperTab.js`を変数に確保する
　→Excelの場合は、分割したやつをJoinする。AccessはシンプルにSelect文

3. 新しいタブ("Target.createTarget")を非表示("hidden:=True")で開く　※必要に応じて表示する設定も可能。以降は、3で開いたタブに対する操作となる
4. `Target.exposeDevToolsProtocol`によるJavaScript操作によるCDP操作権を与える
5. `Runtime.addBinding`で、BiDiのレスポンスを受け取る名称（`sendBidiResponse`）を設定。任意でデバッグ用の名称（`sendDebugMessage`）も設定します
6. `Runtime.evaluate`で、`mapperTab.js`の内容を実行する
7. 6で実行した際の初期化用関数`runMapperInstance`に、3で開いたタブIDを渡して初期化
8. `Runtime.callFunctionOn`の引数に、BiDiコマンドとして載せてJavaScript実行（タブの中の`window.onBidiMessage`を呼ぶ）
9. "Runtime.bindingCalled"を監視し、BiDiの結果なら、RaiseEventでBiDi用のClassへ伝達

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

`hidden` と `Target.exposeDevToolsProtocol` は、CDP では実験的（experimental）な機能です（[Target.pdl の定義](https://github.com/ChromeDevTools/devtools-protocol/blob/7245ec551bc9d450170193456ab3965bb7e716cb/pdl/domains/Target.pdl#L103-L175)）。タブの中の `window.cdp`・`window.sendBidiResponse`・`window.onBidiMessage`・`window.runMapperInstance` は、chromium-bidi の [bidiTab.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiTab/bidiTab.ts#L28-L99) で決まっています。

手順8だけは、使う CDP のメソッドが違います。どれもタブの中の `window.onBidiMessage` を呼ぶ点は同じですが、chromium-bidi と ChromeDriver は `onBidiMessage("…")` という式を組み立てて `Runtime.evaluate` で評価します。このツールは、JSON のダブルクォートやエスケープで式が崩れるのを避けるため、引数として渡せる `Runtime.callFunctionOn` を使っています（[CDPContext.cls の jsEval](https://github.com/Eschamali/StarterWebScrapingKit/blob/ee538fd70e01b7258647c7cb78aca3f8d92b0ce5/src/classes/CDPContext.cls#L655-L675)）。
:::

おわかりいただけただろうか...
BiDiを送っているはずなのに何故か内部ではCDPとして送っていることに...

実は、BiDiをCDPに訳しているのは、Chromium のブラウザ本体（C++ のソース）ではありません。
訳しているのは、BiDiをCDPに翻訳するための[TypeScriptベースのプロジェクト](https://github.com/GoogleChromeLabs/chromium-bidi)（いまの正本は Chromium のソースツリーの `third_party/chromium-bidi`）をビルドした`mapperTab.js`です。Chromium 側の説明にも、「ChromeDriver のための WebDriver BiDi の実装で、mapperTab.js にコンパイルされ、ChromeDriver が Chrome ブラウザへ送り込む」とあります（[README.chromium](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:third_party/chromium-bidi/README.chromium;l=14-16)）。
もう一度言います。Chromium本体のccファイルではなく、jsファイルです。つまり、ビルド済みブラウザのタブで扱える言語です

そうです。
主役は「exe」ではなかったのです。本当の主役は`mapperTab.js`
exeはただ
・BiDi の WebSocket の口を開く
・`mapperTab.js`をタブへ運びセットアップする
・レスポンス内容をパススルーする

という非常にシンプルな役目に降格しました（chromium-bidi の README も、ChromeDriver を「BiDi の WebSocket と HTTP の口を持つ C++ のサーバー」、mapper を「ブラウザの中で動き、BiDi のコマンドを CDP のコマンドに訳すもの」と[分けて説明しています](https://github.com/GoogleChromeLabs/chromium-bidi/blob/c05676b921aa0011c9e124ef91b51eb9dd259e8b/README.md?plain=1#L442-L445)）

今回、Excelで再現したこのフローは[chromium-bidi](https://github.com/GoogleChromeLabs/chromium-bidi)を参考に作り上げました。参考にした Node.js 版の [`src/bidiServer/MapperCdpConnection.ts`](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiServer/MapperCdpConnection.ts) は v17.0.2 まであり、2026年9月に [main から削除されました](https://github.com/GoogleChromeLabs/chromium-bidi/commit/314cf35cfb3de63881fd735b0c3add327d2e2cb2)。
![大まかな手順](/img/story/7.png)

じゃ～WebDriver.exeは、どうなのか...

実は同じフローなのです（違うのは、手順8で `Runtime.evaluate` を使う点くらい）。
なんなら<準備編>の2が、exeの場合は、巨大な配列定数として、`mapperTab.js`がある。完全にExcelと役割として一致しているのです。
ChromeDriver のビルドでは、chromium-bidi の `mapperTab.js` を縮めてから、`kMapperScript` という C++ の文字配列にして埋め込んでいます（[BUILD.gn](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/BUILD.gn;l=76-104)、[embed_bidimapper_in_cpp.py](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/embed_bidimapper_in_cpp.py;l=33-45)、[cpp_source.py](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/cpp_source.py;l=63-71)）。実行時は、この `kMapperScript` をタブへ送り込みます（[session_commands.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/session_commands.cc;l=440-462)）。

![バイナリエディタで検索結果](/img/story/18.png)

著作権表示と共に `<!DOCTYPE html><title>BiDi-CDP Mapper</title>...` という生々しいコードが、数万行にわたってハードコードされていた！
縮めるときも、`Copyright`・`license` を含むコメントは残す設定です（[minify_js.js](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:ui/webui/resources/tools/minify_js.js;l=70-77)）。`<!DOCTYPE html>...` は mapper のタブに表示する画面の HTML で、chromium-bidi の [mapperTabPage.ts](https://github.com/GoogleChromeLabs/chromium-bidi/blob/chromium-bidi-v17.0.2/src/bidiTab/mapperTabPage.ts#L19-L21) にあります。

exe：配列（C++ の文字配列）として埋め込み
Excel：複数セルとして分割配置
Access：1レコード
Puppeteer：npm の依存として chromium-bidi を入れ、直接呼び出す（ただし、タブに `mapperTab.js` を載せるのではなく、[Node.js の中で mapper を動かします](https://github.com/puppeteer/puppeteer/blob/a207faf527ce4305b33a6c0f9706b090c21c115e/packages/puppeteer-core/src/bidi/BidiOverCdp.ts#L59-L66)）

他の有名ツールもWebDriver.exeも自分で組んだVBAも、同じ chromium-bidi の mapper でChromiumをBiDi化していたのだ。
決して無駄なことではないことに安心した瞬間だった。
こうして、[v1.6.0](https://github.com/Eschamali/StarterWebScrapingKit/releases/tag/v1.6.0)（2026年3月）で、外部依存なしでWebDriverBiDiが話せるようになりました。
ドライバー（ChromeDriver）を使って VBA から BiDi を扱う試みは、[note の記事](https://note.com/sele_chan/n/n2b21c7c26ef8)や [WebDriver BiDi for SeleniumVBA](https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA) のように先にありました。ドライバーなしで話すものは、筆者が調べた範囲では見当たらなかったので、VBA 初だと思っています。

![完成](/img/story/12.png)

## Firefoxはどうなのか？

実はこのツール、最近になってFirefoxもついに制御することに成功しました。

さっき、WebDriverBiDiも話せるようになった、とお伝えしました。
しかし、あの時はまだ、かなり密結合な状態で、あくまでも、Chromium専用のBiDiClassとして作っていました。

何故なら、Firefoxは、WebDriverBiDiをネイティブに扱えるからです。Firefox は `--remote-debugging-port` を付けて起動すると、本体に入っている Remote Agent が WebDriver BiDi の口を開きます（[Firefox のソース](https://github.com/mozilla-firefox/firefox/blob/d6fd7a4e69b3c2d4550582eb38d2cdbd9cbad79a/remote/components/RemoteAgent.sys.mjs#L647-L649)、[Firefox のドキュメント](https://firefox-source-docs.mozilla.org/remote/index.html)）。
つまり、このツールの場合は、CDP系Classを経由せずに直でFirefoxに届けないと、一切話せないのです。

WebDriverBiDiというクロスブラウザという規格なのに、Chromiumしか話せない...
そんな矛盾がしばらくは続いてました。

しかし、幾度もの改修工事(RaiseEvents,CDPを真似たClass構成,徹底的な責務分離,OptionClass概念...)を経て、ついに...
メインのコードには、数十行の追記だけで、起動済みFirefoxに限るが、ついにFirefoxもこのツールでも比較的容易に制御できるようになりました

「Chromium で積み上げた設計が、本物の WebDriver BiDi にもそのまま通用した」という、ちょっとした答え合わせになりました


## 将来、ChromiumはWebDriverBiDiをネイティブに扱えるのか？

執筆時点では少なくとも
・Chromiumに、(Firefoxみたいに)WebDriverBiDiをネイティブに扱えるOption起動引数的な物

は存在しませんし、公式やsourceからもそのような示唆は見当たりませんでした。
Chromium の起動引数を定義している [content_switches.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:content/public/common/content_switches.cc;l=592-599) にあるリモートデバッグ用の引数は、`--remote-debugging-pipe`・`--remote-debugging-port` など CDP 用のもので、BiDi という語は出てきません（`content_features.cc`・`chrome_features.cc` も同じ）。
ただ、ChromeDriver のソースには「Mapper をタブの外へ移したら消す」という TODO があり（[chrome_launcher.cc](https://source.chromium.org/chromium/chromium/src/+/b0a44723110d6a7afa3e4ab8c462d05c2c517101:chrome/test/chromedriver/chrome_launcher.cc;l=255-263)）、mapper の置き場所を変える構想自体はあるようです。2026年9月には、chromium-bidi 側も Node.js 版のサーバーを消して、[ChromeDriver 専用になりました](https://github.com/GoogleChromeLabs/chromium-bidi/commit/314cf35cfb3de63881fd735b0c3add327d2e2cb2)。

以降は私の予想になりますが。
・一生、ネイティブ搭載はしない。`mapperTab.js`経由を強いられる

と思っています。
何故そう思うか？理由はシンプルです。
・代替機能がないから

実はこの辺、VBAと似た概念があります。

Microsoftはここんところ、いくつかの製品のサポート終了や廃止を進めています
1. Microsoft Publisher（2026年10月にサポート終了。[告知](https://support.microsoft.com/en-US/publisher/microsoft-publisher-will-no-longer-be-supported-after-october-2026)）
2. VBS（2023年10月に非推奨となり、段階的に削除。[告知](https://learn.microsoft.com/ja-jp/windows/whats-new/deprecated-features)）
3. そしてもちろんIE（2022年6月にサポート終了。[告知](https://learn.microsoft.com/ja-jp/lifecycle/announcements/internet-explorer-11-end-of-support)）

代表に挙げた3つはどれも代替品があります
1. Microsoft Create、Word や PowerPointにあるテンプレート系等
2. ps1（Microsoft は PowerShell と JavaScript への移行を[勧めています](https://techcommunity.microsoft.com/blog/windows-itpro-blog/vbscript-deprecation-timelines-and-next-steps/4148301)）
3. Edge（Microsoft が IE からの移行先として勧めているのは、Chromium ベースの Edge）

では、VBAはどうか？
ないんです。

何故か？では、下記の代替機能は他のプログラミング言語にありますか？特にExcelVBAはどうです？
・環境構築不要(Officeインストールというインストールexeを実行で、Office開いてalt + F11でStart)
　→他言語の場合は、追加でvscodeが要る

・ワークシートがGUIエディター
　→他言語の場合は、xaml等と言ったコードページのUI構成言語の知識を要求される

・データとプログラムが一心同体(ワークシート/Accessテーブル)
　→他言語の場合は、SQLiteやini,csvといった外部ファイル化することになり、意図せずファイル移動/削除なんかしたら、もう動かなくなることも
　　責務分離としては美しいので、ここは賛否両論ありそう。

・ファイル配布しやすさの異常さ
　→他言語の場合は、ZIPとして見た目上、1ファイルとしてやる必要があり、実際に実行する際は、結局ちらばるし、Runtime(python,node.js)がなかったら詰み
　　まぁ、Officeもなかったらそれも詰みではあるが、「会社」という括りなら、相殺される

・追加申請不要(業務遂行上必要なファイルとして情シスの検問をPASS)
　→他言語の場合は、要申請になるし、待たされた挙句拒否されることも。
　　まぁ、これは個人PCにとっては意味のない記述だが、「会社」という括りなら、価値あるメリット

どうです？この5つを全て叶えるプログラミング言語、思いつきます？🙃🙃思いつかないと思います。

意外かもしれませんが、
銀行、証券、工場、官公庁……。
世界中の名だたる大企業の「絶対に止めてはいけない日次処理」が、誰が書いたかも分からないスパゲッティVBAで動いています。
Microsoftがもし「明日からVBAを廃止します」と言った瞬間、経済が止まるほどです。

CDPも実は同じです。思いつきますか？CDPの代替機能。
ないです。だってGoogle（Chromium）自身が育ててきた、プログラミング言語ならぬ、プロトコルなのですから。
Chrome チームも、「WebDriver BiDi は CDP を置き換えない。Chromium 系のブラウザは、デバッグのために CDP を使い続ける」と[書いています](https://developer.chrome.com/blog/webdriver-bidi)。

唯一近いWebDriverBiDi、でもこれはあくまでも、自動化という括りまでの制定です。
CDPはデバッグ兼自動化。BiDiは自動化だけ。
CDP の[公式ページ](https://chromedevtools.github.io/devtools-protocol/)は CDP を「デバッグと自動化の対象を、デバッグと自動化のクライアントとつなぐプロトコル」と、W3C は WebDriver を「ユーザーエージェントを調べて操作するための[遠隔操作のインターフェース](https://www.w3.org/TR/2018/REC-webdriver1-20180605/)」と書いています。

どっちがカバー範囲が大きいかは明らかですね。
ただ、それだけのことです。


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

## 次に読む

- [アーキテクチャ](/concepts/architecture) — mapper が載る層の位置づけ
- [CDP と BiDi](/concepts/cdp-vs-bidi) — どちらを使うか
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol) — `ExecuteBiDi` / BiDi+
