---
description: WebDriver なしでWebDriverBiDiが扱えるとわかった経緯。公式ドライバーに頼らない BiDi 対応の登場秘話です。
---

# なぜ外部依存なしで、ChromiumにWebDriverBiDiが扱えるとわかったのか


![海外フォーラムでの議論](/img/story/1.png)
すべては[この議論](https://github.com/GCuser99/SeleniumVBA/discussions/164#discussioncomment-15988897])から始まりました。

> Well, I think we had this discussion earlier regarding BiDi. 
> And when it comes to companies not allowing local software installations: if Chrome is not being allowed to install, then CDP is worthless anyway. So even in a scenario of restricted environments I assume CDP support not to be a breakthrough.
さて、この話は以前にもBiDiについて話したと思います。CDPはBiDiほど将来性に欠けています。
そして、企業がローカルソフトウェアのインストールを許可しない場合、Chromeのインストールが許可されていなければ、CDPは結局無意味です。ですので、制限された環境のシナリオでも、CDPサポートは画期的なものではないと考えます。

この頃になると、CDPに対する知識は随分と溜まってきました。
なので、すぐに気付きました。半分あっていて、半分間違ってる

CDPは、Chromeブランドではなく、Chromiumブランドです。よって、業務PCにプリインストールされているEdgeもCDP対応するということになります。
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
CDP:Google
WebDriverBiDi:W3C

W3Cは簡単に言うと、大手のブラウザメーカーたちが集まって共通のコマンド名でうちらのブラウザを操る世界標準を作るグループみたいなもんです。

全ての始まりは、GoogleでもなくMicrosoftではなく、Seleniumです。
そうです。VBAでは標準となってる`SeleniumVBA`もSeleniumというメーカーの上で成り立っています

最初の自動化はどうしてたのかというと...
なんと！ネットワークの書き換えです🫨

要するに、間にプロキシサーバーを挟んでその内部で受信したhtmlにテストツールのjsを仕込む

という今見たらもはやウイルスチックなことをしてました。でも当時はこれしか選択肢がなかったのです🙃
しかし、当然この仕組みは限界があります。
あくまでも、主体はJavaScriptなので、JavaScriptのルール違反になるようなことをしようとすると当然動きません。
でも、そのルール違反的なことをしないと、テストもままにならないのも事実...

そして時が経ちようやくここでおなじみのWebDriverが登場しました...いや厳密には「WebDriver」という製品名が来ました

でも内部ロジックとしては今と同じように
・各ブラウザごとの操作内容を隠ぺい/共通化
しています。

例えばこのツールのIntroとして出してるIEならCOM
別ブランチ(ex-Firefox)で密かに対応してるFirefoxなら、当時はXPCOM
今主流のChromeなら、当時はNPAPI

当然こんなバラバラ相手に1つ1つ学習してたらキリがありません。
それをWebDriverという製品が解決してくれました

そしてSeleniumは気付きました
・うちのJavaScriptハックよりもこのWebDriverのほうがええやん🙃

こうしていい感じに統合されて、ここでようやく、規格としてのWebDriverが登場しました。
複数ブラウザ相手なのに覚えるAPIはWebDriverという巨大なでも1種類だけ
当時は画期的でした。

しかし、Webの進化は異常に進みました。
ちょっと思い出してみてください。

昔、ブラウザでゲームと言えば
・flashPlayer
・Java

みたいな今でいう拡張機能を入れる必要がありました。
「これを使うには、flashPlayをインストール...」
一体何回見たことやら...

また、当時のYoutubeも実は動画部分をflashPlayerに依存してました

でも今のブラウザはどうです？
・URLを共有するだけで動画もゲームもすぐにできる/見れる

理由は単純です。それらをつかさどる機能がブラウザに丸ごと次々と標準搭載されたからです
・WebGLAPI
・WebAUDIOAPI
・WebAssemble
・CANVAS
...AndMore

こういった技術の結集が高速に実装され、リンク共有だけで素晴らしい体験ができました。

...しかし、この進化に追いつけてないのが1つだけありました。
はい、WebDriverです。

WebDriverの基本は、送って応答待って回答を得るという同期スタイルです。(技術的にいうとhttpといった同期やり取り)
しかしさきほどの進化の影響でこの仕組みに限界が出てきました
・ページは読み込み終わったのに、ページに表示されてる内容が書き換わる
　→検索ボタンを押したら、まるごとページ読み見直しではなく、部分的に読み直す

・inputタグの動的生成対応
　→初期状態にはなく、ボタン押下することで初めて出現するタイプ。検索では見つからない

・ネットワークの監視
　→非同期処理によるデータ内容に応じた処理

いずれもこれは、SPA（SinglePageApplication）サイトによる「内部で起こった出来事」挙動です。技術が発展したからこそ登場したサイトです。

しかし実はこれら3つとも解決できる手段はあったっちゃありました。
はい、このツールの主題であるCDPです。

元々はデバッグ操作用としてGoogle独自に実装した2011年発のコマンドです。以外にも、WebDriverよりも後発です😳
ChromiumもCDPもGoogleが作ってます。当然、Webの進化とともに、CDPも同時に進化していきました。
そして進化していくうえで気づいてしまいました
>あれ？これブラウザ自動化にも使えるじゃん

これが、あのPuppeterの登場です。

しかし、名称の通り、あくまでもChromium(Chrome/Edge)製でしか通じないプロトコルです。
このままでは、またあの
・各ブラウザ相手1つ1つ学習してたら鬱になる

が戻ってしまいます

こうして、現在のWebDriver BiDiの制定にいたりました。

## ChromiumをWebDriverBiDiとして話す方法

さて、お待たせしました。ようやく、本題です。
大まかな歴史をつかんだところで、このツールはどうやって、ChromiumをBiDiとして喋っているのか...

### WebDriver.exeルート

とその前に、一般的にWebDriver.exe経由でChromiumに、このWebDriverBiDiを扱うにはどうするのか？
ざっくりフローとしてはこうです
1. DLしたWebDriverに対して、「--port=9222」オプションをつけて起動する。
2. 空いたPortに向かってcapabilitiesに「webSocketUrl: true」と送る
3. すると内部で`session.new`が発火し、WebDriverBiDiが有効。`session.NewResult`が返る
4. レスポンスに`webSocketUrl`にWebSocketURLが来るのでそこにWebSocket接続を行う
5. あとはBiDiのルールとしてコマンドのやり取りを、WebSocket上で行う

WebDriver.exeで、BiDiを扱う詳細なフローは[ここで](https://note.com/sele_chan/n/n2b21c7c26ef8)確認できます。

### このツール

では、このツールの場合はどうなのか...
大きく2フェーズあります。
こうなります。

<準備編>
1. npm等で`mappertab.js`を入手
2. 中身をExcel/Accessに保存
　→Excelの場合は、セル内に保持できる総文字数が、32,767 文字なので、百万近くある`mapperTab.js`文字数ををその文字数として分割保存してます。
　　Accessは、1レコードに収める「長いテキスト」型があるので、特に難なく入れれた

<実際に動かす>
1. BiDiとして使用を開始する`Sub`を呼び、Chromiumのデバック(Port/Pipe)に接続
2. 準備編で保存した`mapperTab.js`を変数に確保する
　→Excelの場合は、分割したやつをJoinする。AccessはシンプルにSelect文

3. 新しいタブ("Target.createTarget")を非表示("hidden:=True")で開く　※必要に応じで表示する設定も可能。以降は、3で開いたタブに対する操作となる
4. `Target.exposeDevToolsProtocol`によるJavaScript操作によるCDP操作権を与える
5. `Runtime.addBinding`で、BiDiのレスポンスを受け取る名称を設定。任意でデバッグ用の名称も設定します
6. `Runtime.evaluate`で、`mapperTab.js`の内容を実行する
7. 6で実行した際の初期化用関数`runMapperInstance`に、3で開いたタブIDを渡して初期化
8. `Runtime.callFunctionOn`の引数に、BiDiコマンドとして載せてJavaScript実行
9. "Runtime.bindingCalled"を監視し、BiDiの結果なら、RaiseEventでBiDi用のClassへ伝達

おわかりいただけただろうか...
BiDiを送っているはずなのに何故か内部ではCDPとして送っていることに...

実は、ChromiumSource(cファイル等)そのものには、BiDiを話すロジックは備わっていません。
あるのは、BiDiをCDPに翻訳するための[TypeScriptベースのプロジェクト](https://github.com/GoogleChromeLabs/chromium-bidi)をビルドした`mapperTab.js`があるだけです。
もう一度言います。Chromium本体ccファイルではなく、jsファイルです。つまり、ビルド済みブラウザのタブで扱える言語です

そうです。
主役は「exe」ではなかったのです。本当の主役は`mapperTab.js`
exeはただ
・`mapperTab.js`をタブへ運びセットアップする
・レスポンス内容をパススルーする

という非常にシンプルな役目に降格しました

今回、Excelで再現したこのフローは[chromium-bidi](https://github.com/GoogleChromeLabs/chromium-bidi)を参考に作り上げました。
![大まかな手順](/img/story/7.png)

じゃ～WebDriver.exeは、どうなのか...

実は全く同じフローなのです。
なんなら<準備編>の2が、exeの場合は、巨大な配列定数として、`mapperTab.js`がある。完全にExcelと役割として一致しているのです。

![バイナリエディタで検索結果](/img/story/18.png)

著作権表示と共に `<!DOCTYPE html><title>BiDi-CDP Mapper</title>...` という生々しいコードが、数万行にわたってハードコードされていた！

exe：配列として分割配置
Excel：複数セルとして分割配置
Access：1レコード
Puppeter： 依存関係の設定により、chromium-bidi リポジトリから直接呼び出される

他の有名ツールもWebDriver.exeも自分で組んだVBAも、全く同じフローでChromiumをBiDi化していたのだ。
決して無駄なことではないことに安心した瞬間だった。
こうして、v1.6.0で、VBA初の外部依存なしでWebDriverBiDiが話せるようになりました

![完成](/img/story/12.png)

## Firefoxはどうなのか？

実はこのツール、最近になってFirefoxもついに制御することに成功しました。

さっき、WebDriverBiDiも話せるようになったっっとお伝えしました。
しかし、あの時はまだ、かなり密結合な状態で、あくまでも、Chromium専用のBiDiClassとして作っていました。

何故なら、Firefoxは、WebDriverBiDiをネイティブに扱えるからです。
つまり、このツールの場合は、CDP系Classを経由せずに直でFirefoxに届けないと、一切話せないのです。

WebDriverBiDiというクロスブラウザという規格なのに、Chromiumしか話せない...
そんな矛盾がしばらくは続いてました。

しかし、幾度な改修工事(RaiseEvents,CDPを真似たClass構成,徹底的な責務分離,OptionClass概念...)を得て、ついに...
メインのコードには、数十行の追記だけで、起動済みFirefoxに限るが、ついにFirefoxもこのツールでも比較的用意に制御できるようになりました

「Chromium で積み上げた設計が、本物の WebDriver BiDi にもそのまま通用した」という、ちょっとした答え合わせになりました


## 将来、ChromiumはWebDriverBiDiをネイティブに扱えるのか？

執筆時点では少なくとも
・Chromiumに、(Firefoxみたいに)WebDriverBiDiをネイティブに扱えるOption起動引数的な物

は存在しませんし、公式やsourceからもそのような示唆は見当たりませんでした。

以降は私の予想になりますが。
・一生、ネイティブ搭載はしない。`mapperTab.js`経由を強いられる

と思っています。
何故そう思うか？理由はシンプルです。
・代替機能がないから

実はこの辺、VBAと似た概念があります。

Microsoftはここんところ、いくつかの製品サポートを終了させています
1. Microsoft Publisher
2. VBS
3. そしてもちろんIE

代表に挙げた3つはどれも代替品があります
1. Microsoft Create、Word や PowerPointにあるテンプレート系等
2. ps1
3. Chrome(Edge)

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
ないです。だってGoogle自身が作ったプログラミング言語ならぬ、プロトコルなのですから。

唯一近いWebDriverBiDi、でもこれはあくまでも、自動化という括りまでの制定です。
CDPはデバック兼自動化。BiDiは自動化だけ。

どっちがカバー範囲が大きいかは明らかですね。
ただ、それだけのことです。


## 次に読む

- [アーキテクチャ](/concepts/architecture) — mapper が載る層の位置づけ
- [CDP と BiDi](/concepts/cdp-vs-bidi) — どちらを使うか
- [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol) — `ExecuteBiDi` / BiDi+
