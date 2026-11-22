---
description: Web ページ（ブラウザ）とローカルのアプリが、くっついたり離れたりしてきた歴史です。IE・HTA の時代から WebView2 と CDP の時代までを、双方向のやり取りの仕組みに絞って振り返ります。
---

# Web とローカルの双方向やり取り、30年の輪廻

::: tip 秘話
〜くっついて、離れて、またくっつく〜
:::

## はじめに

一般的な Web スクレイピングは、こちらから命令を投げて、その結果が返ってくるまで待つ、という流れでした。しかし Web の進化とともに、その流れだけでは限界が見えてきます。

そこで近年のツールでは、ブラウザの中で起きた出来事（Network 系、Page 系など）を、イベントとしてブラウザから一方的に送ってくる仕組みができました。それが CDP / WebDriver BiDi です。

たいていはこれで十分なのですが、さらにこんな要望も出てきました。

- Web ページの中の出来事が任意の条件に当てはまったら、それもイベントとして送ってほしい🥺

ブラウザから一方的に届く仕組みはあっても、届くのはブラウザ側があらかじめ決めたイベントだけです。新しいイベントを足すかどうかは、CDP なら Google、BiDi なら W3C に委ねられます。つまりこの要望は、「イベントを自分で作りたい」というものです。

この要望を聞いて、ふと思いました。

> あれ、それって、IE を自分のアプリに埋め込んで双方向にやり取りしていた、昔のあの仕組みと同じじゃね？😳😳

今回は、この「Web ページとローカルのアプリの双方向やり取り」に絞って、その歴史を語ります。VBA から実際に使う方法は、[ブラウザと VBA の双方向やり取り](/guides/two-way)にまとめています。

---

## 第1幕: 壁がなかった時代（1990年代後半〜2000年代前半）🏡

### IE を自分のアプリに埋め込む

最初の Web ページとローカルのアプリの連携術は、おそらく IE です。

1996年、Microsoft は Web ページに Windows のネイティブの技術（COM/OLE）を埋め込める ActiveX と、Web 向けのスクリプト言語 VBScript を世に出しました。IE そのものも、WebBrowser コントロールとして、自分のアプリに部品のように埋め込めました。

埋め込んだ IE とホストのアプリは、両方向に話せました。.NET の WebBrowser コントロールなら、ホストのオブジェクトを `ObjectForScripting` に渡すと、ページの JavaScript から `window.external` で呼べます。逆向きは `Document.InvokeScript` で、ページの関数を呼べます。Microsoft の手順書の題も、そのものずばり「DHTML コードとクライアント アプリケーション コード間の双方向の通信を実装する」です。

VBA も例外ではありません。`InternetExplorer.Application` や UserForm の WebBrowser コントロールを `WithEvents` で持てば、ページ上のクリックまで VBA のイベントプロシージャで受け取れました。IE と VBA は、同じ COM の世界に住んでいたからです（具体的な書き方は[ブラウザと VBA の双方向やり取り](/guides/two-way)の「はじめに」にあります）。

デスクトップそのものにも Web が入り込みました。IE4 と Windows 98 のアクティブデスクトップは、Windows のデスクトップと IE を統合し、ローカルのファイルも、イントラネットも、Web も、同じ見た目で扱えるようにしました。

### 究極形、HTA

1999年には IE5 とともに、HTA（HTML アプリケーション、`mshta.exe`）が登場します。HTML と JavaScript（や VBScript）だけで、ローカルのアプリが作れる仕組みです。

Microsoft の解説によると、HTA は「完全に信頼されたアプリ」として、IE が Web ページには決して許さない操作をこなします。ゾーンのセキュリティ設定に関係なく ActiveX コントロールも動かせたので、`new ActiveXObject("Excel.Application")` と書けば、HTA の JavaScript から Excel を立ち上げて操ることもできました。

メモ帳で数行書いて拡張子を `.hta` にするだけで、OS に最初から入っている IE を借りて動くデスクトップアプリができあがる。今思うと、なかなか自由な時代です😳

---

## 第2幕: 壁ができた時代（2000年代〜2010年代）🚧

### 便利すぎた代償

壁がないということは、悪意のあるものも素通りできるということです。

2000年5月には、「ILOVEYOU」という題のメールに `LOVE-LETTER-FOR-YOU.TXT.vbs` を添付したワームが、世界中で流行しました。開いてしまうと、Outlook のアドレス帳の宛先へ自分自身を送り、手元のファイルを消しては自分自身に置き換えて回ります。中身は VBScript でした。

それでも、Web の技術でローカルを動かす仕組みは作られ続けます。Windows Vista と 7 のサイドバーに並べるガジェットも、その 1 つです。しかし 2012年7月、Microsoft はセキュリティ アドバイザリ 2719662 で、サイドバーとガジェットを丸ごと無効にする Fix it を公開します。悪意のあるガジェットや、正規のガジェットの脆弱性を突かれると、ユーザーの権限で任意のコードを実行されるおそれがある、という理由です。その後、ガジェットの配布ページは「ガジェットは提供を終了しました」という案内に変わり、次の説明が載りました。

- Windows 7 と Vista のサイドバーの仕組みには、深刻な脆弱性がある
- 新しい Windows では、この機能をなくした

### ブラウザは別棟へ

一方、ブラウザの側は「ページにローカルを触らせない」方向へ進みます。

2008年に登場した Google Chrome は、ページを描くプロセスをサンドボックスに閉じ込めました。OS のセキュリティの仕組みを使い、ページのコードがコンピューターに永続的な変更を加えたり、機密の情報にアクセスしたりできないようにするものです。

Microsoft も 2015年、新しいブラウザ Microsoft Edge（当時は EdgeHTML）で、ActiveX と VBScript に別れを告げます。

壁の向こうへ話しかけたい開発者は、抜け道を掘りました。ローカルのアプリが `localhost` に小さな Web サーバーを立ててページから呼ばせる方法や、独自の URL スキームでアプリを起動させる方法です。Chrome の拡張機能には、ローカルのアプリとやり取りするための公式の通り道として、ネイティブ メッセージングも用意されました。Chrome がローカルのアプリを別のプロセスで起動し、標準入出力で JSON のメッセージを送り合う仕組みです。

### それでも IE を組み込むアプリは止まらない

壁ができても、IE を組み込んだアプリは止まりませんでした。

2022年6月15日、IE11 のデスクトップ アプリはサポートを終えます（一部の Windows 10）。ただし、その FAQ には大事な一文があります。IE の土台の MSHTML（Trident）エンジンは引き続きサポートされ、それに頼る WebBrowser コントロール（WebOC）と MSHTA も動き続ける、というものです。IE の画面がなくなっても、組み込み用の IE は生き残りました。

Excel でおなじみの Power Query が、IE の終了後も動いているのはこのためです。Web からデータを取る `Web.Page` 関数は、IE の上に作られています。ただし画面からは、WebView2 を使う `Web.BrowserContents` への置き換えが進んでいます。Power BI Desktop の OAuth の認証画面も、2021年10月から既定で WebView2 になりました。

IE に頼る業務システムも、特に日本ではしぶとく残りました。IE のサポート終了直後（2022年6〜7月）の国内の調査では、業務で「IE を使っている」が 12％、「Edge の IE モードを使っている」が 35％でした。

日本の行政のシステムも、少しずつモダンなブラウザへ移っています。e-Tax は 2020年5月に Chrome と新しい Edge へ一部対応しましたが、使うには拡張機能「e-Tax受付システムAP」と、「事前準備セットアップ」というローカルのインストーラーが要ります。マイナンバーカードを読み取る公的個人認証（JPKI）の利用者クライアントソフトも、Edge / Chrome 版はブラウザの拡張機能と組み合わせて使います。ページから直接ローカルに触れなくなった分、拡張機能とローカルのソフトが手を組む形です。

---

## 第3幕: 組み込み IE の正統進化、WebView2（2020年〜）🌉

### 組み込み IE の引退

IE11 のデスクトップ アプリは 2022年に終わりました。残る Edge の IE モードも、Microsoft が 2026年9月に「少なくとも 2029年末までサポートする」と改めて周知し、早めの移行を呼びかけています。2025年10月には、攻撃への悪用を受けて、IE モードの使い方に制限も加わりました。

組み込み用の IE も、活躍の場が減っています。

| 製品 | 内容 | 時期 |
| --- | --- | --- |
| Access | IE ベースの旧 Web ブラウザー コントロールを廃止する。Web の内容を読み込まず、表示もしなくなる | 2027年5月18日（バージョン 2705）の予定 |
| Outlook | MSHTML（Trident）を使う Outlook Today とフォルダー ホーム ページを非推奨にし、その後、完全に無効にする | 2026年8月（バージョン 2608）に非推奨、2027年8月（バージョン 2708）に無効 |
| Power Query | IE ベースの `Web.Page` を、画面からは WebView2 ベースの `Web.BrowserContents` に置き換える | 進行中 |

Outlook の告知は、この非推奨を「Windows が MSHTML（Trident）エンジンを非推奨にしたことに続く」ものだと説明しています。

### WebView2 の登場

組み込み用の IE の正統な後継が WebView2 です。2020年10月19日、Win32 C/C++ 向けに一般提供が始まりました。

WebView2 には、最初のインターフェイス `ICoreWebView2`（SDK 0.9.430 で登場）の時点で、ページとホストが話す仕組みがそろっていました。

- `AddHostObjectToScript`: ホストのオブジェクトを、ページの JavaScript から呼べる。IE の `window.external` の生まれ変わりです
- `PostWebMessageAsJson` / `WebMessageReceived` と、ページ側の `chrome.webview.postMessage`: 文字列や JSON のメッセージを、両方向に送り合う

ふつうの軽いデータなら、これで十分でした。ところが、今の Chromium は重い処理にも耐えられるよう、劇的に進化しています（ブラウザでゲームやシミュレーションが動くほど😳）。そうなると、画像や PDF のような大きなデータも扱いたくなってきます。

### パフォーマンスのお話

ここで、上の 2 つの仕組みが壁になります。WebView2 の要望を受け付ける GitHub のリポジトリには、次のような声が寄せられていました。

- 2019年12月: JSON と文字列でデータを運ぶのは、とにかく遅すぎる。印刷用に 1 ページ 100MB のデータを JavaScript から C++ に渡したいが、base64 の文字列にして送るしかなく、遅くて反応も悪い
- 2021年2月: JSON に変えて運び、受け取った側で解析し直す代わりに、データを直接渡せれば大きく速くなる

極論を言うと、VBA でみんなが一度は踏んで Excel をフリーズさせる、あの悪名高い `A = A & B`（文字列結合）と構造がそっくりなんですよね。`A = A & B` は、つなぐたびに新しい文字列を丸ごと作り直します。メッセージも、送るたびにデータを丸ごと文字列にして運び、受け取った側で作り直します。データが大きいほど、この「丸ごと」が効いてきます。

::: info
WebView2 の中のメッセージの通り道は、ソースが公開されていません。実際に何回コピーしているかまでは確かめられず、ここに書いたのは、要望を出した人たちの声から分かる範囲です。
:::

### SharedBuffer の登場

そこで追加されたのが SharedBuffer です（SDK 1.0.1661.34 で正式に入りました）。仕様の背景には、まさに先ほどの 2 つの要望が挙げられています。OS の共有メモリを使い、ホストとページが同じメモリを読み書きします。

VBA でいうなら、`A = A & B` をやめて、確保済みのバッファに `Mid$(buf, pos, len) = B` で書き込むような発想です。このツールでの使い方は、[ブラウザと VBA の双方向やり取り](/guides/two-way)の「SharedBuffer」にあります。

---

## 第4幕: 普通のブラウザでも…🌍

普通のブラウザ（Chrome や Edge）は、基本的に Web ページの中で完結します。IE の反省から、サンドボックスを第一に作られているからです。だからこそ、安心してネットサーフィンができます。

でも、ある場面ではこれが厄介です。そうです、テストの自動化です。そして実は、普通のブラウザにも、ある面でこうした双方向のやり取りの仕組みが登場しちゃいました。

こうした Web ページとサードパーティ製アプリの連携には、これまで IE や WebView2 のような専用のライブラリが必要で、OS も Windows に縛られていました。しかし、今のクロスプラットフォームな Chromium は、デバッグ用の引数（`--remote-debugging-port` / `--remote-debugging-pipe`）や `chrome://inspect/#remote-debugging`（Edge なら `edge://inspect/#remote-debugging`）を使うことで、理論上、Web ページとサードパーティ製アプリの連携ができてしまいます。そのうち「ページからアプリへ」の通り道が、`Runtime.addBinding` です。

ただ、`Runtime.addBinding` は最初からあったわけではありません。

### `console.debug` で我慢していた時代

`Runtime.addBinding` が無かったころ、ツールは `console.debug` を通り道の代わりにしていました。たとえば Puppeteer の v1.0.0 では、`page.exposeFunction` で公開した関数が呼ばれると、ページ側で `console.debug('driver:page-binding', ...)` を実行します。ツール側は、コンソールの出力を知らせる `Runtime.consoleAPICalled` を受け取り、`'driver:page-binding'` の印が付いたものを、関数の呼び出しとして扱っていました。

ここにも、遅さの壁がありました。

### `Runtime.addBinding` の登場

2018年5月、V8 の DevTools 担当の Alexey Kozyatinskiy が、V8 に `Runtime.addBinding` と `Runtime.bindingCalled` を追加します（コミットの題では `Runtime.installBinding`）。コミットメッセージには、追加した理由が次のように書かれています。

- 多くのクライアントが、ページからプロトコルのクライアントへの通信路として `console.debug` を使っている。`console.debug` は少し遅く、この使い方を想定して作られていない
- 追加した関数は、マイクロベンチマークで `console.debug` より約 4.6 倍速い

同じ年の6月には、同じ作者が Puppeteer の `page.exposeFunction` を `console.debug` から `Runtime.addBinding` へ切り替えています（コミットメッセージは「New way is faster and cleaner.」）。

| 時期 | 出来事 | 出典 |
| --- | --- | --- |
| 2018年まで | Puppeteer などが `console.debug` と `Runtime.consoleAPICalled` で、ページからの合図を受け取る | [Puppeteer v1.0.0 の `lib/Page.js`](https://github.com/puppeteer/puppeteer/blob/v1.0.0/lib/Page.js) |
| 2018年5月31日（V8 6.9 = Chrome 69） | V8 に `Runtime.addBinding` / `Runtime.removeBinding` / `Runtime.bindingCalled` が入る | [V8 のコミット](https://chromium.googlesource.com/v8/v8/+/5cb11a17cdd1762f52ef5e46a5c963e33c5d5b06) |
| 2018年6月5日 | `executionContextId` を指定して、1つのコンテキストだけに追加できるようになる | [V8 のコミット](https://chromium.googlesource.com/v8/v8/+/e63f35bcaca4dcac5596d9707791a0c251ebea87) |
| 2018年6月18日 | Puppeteer の `page.exposeFunction` が `Runtime.addBinding` に切り替わる。Chromium 69 を同梱した v1.6.0 で公開 | [Puppeteer のコミット](https://github.com/puppeteer/puppeteer/commit/b20cde67c63a6344321c2fc40ef2d3cbf8c3ae24) / [v1.6.0](https://github.com/puppeteer/puppeteer/releases/tag/v1.6.0) |
| その後 | chromium-bidi が、Mapper のタブからの WebDriver BiDi の応答を `Runtime.addBinding`（`sendBidiResponse`）で受け取る。このツールの BiDi も同じ形 | [chromium-bidi の `bidiTab.ts`](https://github.com/GoogleChromeLabs/chromium-bidi/blob/bb98fab480403c61f322320f2ebd06c617ed0b90/src/bidiTab/bidiTab.ts) |

もともとは、自動化ツールのための「`console.debug` より速い通り道」として入った機能です。それがデバッグ用の接続口と組み合わさることで、Windows 以外も含めた普通のブラウザで、ページからサードパーティ製アプリへ合図を送れるようになりました。

使い方はシンプルです。どんな名前で登録し、どんな条件でその名前の関数を呼ぶかを、JavaScript で仕込んでおきます。あとは条件に引っかかって呼ばれれば `Runtime.bindingCalled` が飛んでくるので、その中身をアプリが読み取るだけです。渡せるのは `payload` という文字列 1 つだけですが、JSON の文字列にすれば、いくらでも項目を足せます。逆に言えば、大きなデータには向きません。あくまで、軽いデータのやり取り用です。

### 一緒やん🙃

軽いデータ……そうです。先ほど挙げた、WebView2 のやり取りの方法を思い出してください。

- `AddHostObjectToScript`
- `WebMessage` × `chrome.webview.postMessage`

役割はそっくりです。特に `WebMessage` × `chrome.webview.postMessage` は、ページで呼べば、ホストに COM のコールバックとして非同期で届きます。`Runtime.addBinding` で登録した名前で呼べば、`Runtime.bindingCalled` というイベントとして非同期で届きます。完全一致です😳

| | WebView2 | デバッグ用に起動した Chromium |
| --- | --- | --- |
| ページ → アプリ | `chrome.webview.postMessage` → `WebMessageReceived` | 登録した関数を呼ぶ → `Runtime.bindingCalled` |
| アプリ → ページ | `PostWebMessageAsJson` | `Runtime.evaluate` |
| 届き方 | COM のコールバック（非同期） | CDP のイベント（非同期） |

つまり、次の 2 つはイコールの関係です。やろうと思えば作れてしまう、理論上の関係です。

- WebView2 で、`WebMessage` × `chrome.webview.postMessage` を使った軽量アプリ
- デバッグ用に起動した Chromium で、`Runtime.addBinding` × `Runtime.bindingCalled` を使った軽量アプリ

ただ、VBA から見れば、`AddHostObjectToScript` のほうが直感的で、返事も受け取りやすいですね。

まったく違うところで生まれたメソッドが、掘り下げると同じ考え方にたどり着く。面白いですね。

### それでも大容量は苦手

とはいえ、大きなデータは扱えません。理論上は base64 にすれば送れなくはないですが（約 1.3 倍に膨らみます）、CDP の通り道には、送受信の容量の上限があります。Chromium のソースを見ると、次のとおりです。

| 通り道 | ブラウザ → アプリ | アプリ → ブラウザ |
| --- | --- | --- |
| WebSocket（`--remote-debugging-port`） | 256MB（送りきれていないデータの上限） | 100MB |
| パイプ（`--remote-debugging-pipe`） | 上限の定数なし（64KB ずつ書き出す） | 100MB |

意外にも、VBA の可変長の String に入る約 20 億（2^31）文字より、はるかに小さいんですよね。スクリーンショットのように、大きなデータを base64 で返す CDP のコマンドを使うときは、少し注意したほうがよさそうです。

---

## 第5幕: 壁が閉まりかけて、また開く（2025年〜）🤖

2025年、普段使いの Chrome のリモート デバッグを悪用して、Cookie を抜き取る攻撃が増えました。そこで Chrome 136 からは、既定の `User Data` に対しては `--remote-debugging-port` / `--remote-debugging-pipe` が無視されるようになります。使うには、`--user-data-dir` で別のフォルダを指定します。壁が、また閉まりかけたわけです。

ところが同じ年の暮れ、AI のコーディング エージェントにブラウザを操作させる時代がやってきます。Chrome 144 では、`chrome://inspect/#remote-debugging` で有効にすれば、ユーザーが許可したうえで、普段使いのブラウザにも接続できるようになりました。人が「許可」を押したときだけ開く関所です。VBA でいう「コンテンツの有効化」バーの再来ですね🙃

| 時期 | 出来事 | 出典 |
| --- | --- | --- |
| 2025年（Chrome 136） | 既定の `User Data` に対しては、`--remote-debugging-port` / `--remote-debugging-pipe` が無視されるようになる。使うには `--user-data-dir` で別のフォルダを指定する | [Chrome のブログ](https://developer.chrome.com/blog/remote-debugging-port?hl=ja) |
| 2025年12月（Chrome 144） | `chrome://inspect/#remote-debugging` で有効にすると、普段使いのブラウザにも、ユーザーが許可したうえで接続できるようになる | [Chrome のブログ](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja) |

---

## 〆: 25年かけた「名前を変えただけの輪廻転生」

Web ページとローカルのアプリの連携は、誰でも触れられる IE や HTA として登場しました。しかし、セキュリティの問題から、そうした双方向のやり取りは封じられていきます。

それでも、次の要求は今も消えません。

- テストを自動化したい🥺
- HTML / CSS / JS でサクッときれいな画面を作って、ローカルを動かしたい！

後者には、正統な後継として IE → WebView2、HTA → Electron が現れ、ローカルと連携する仕組みが戻ってきました。前者にも、条件付き（デバッグ用の引数で起動するか、`inspect/#remote-debugging` で許可する）で、ローカルと連携する仕組みが来ました。

| | 1990年代後半（Microsoft の発明） | 2020年代（今の技術） |
| --- | --- | --- |
| アプリにブラウザの部品を埋め込む | IE（WebBrowser コントロール） | WebView2 |
| HTML と JS で、ローカルを触るアプリを作る | HTA（`mshta.exe`） | Electron |

……やってること、1ミリも変わってない！！！🙃🙃🙃

しかも「軽さ」は、むしろ HTA のほうが上だった、という皮肉つきです。

- **HTA**: OS に最初から入っている IE のエンジンを借りて動くので、メモ帳で数行書くだけの、数 KB のアプリが作れた
- **Electron**: Chromium と Node.js を丸ごと抱え込むので、Hello World だけでも大きなアプリになる
- **Tauri**: 「Electron は重すぎじゃね？ Windows なら最初から入っている WebView2 を借りようぜ！」 → それ、HTA への先祖返りやないか！😂

名前を変えてやり直すことになった理由も、結局は同じです。

1. HTA も IE も、自由で便利すぎたせいで、悪用されまくった💀
2. Microsoft がセキュリティのために、いったん封印した
3. でも、「HTML / CSS / JS でサクッときれいな画面を作って、ローカルを動かしたい！」という人類の欲望は消えない
4. 名前と中身（Edge / Chromium）を新しくして、もう一度同じものを作り直す

技術の進化は、右肩上がりに見えて、実は何かの無限ループに捕らわれているのかもしれません。もうこれ、IT 業界最大のジョークですよね。

**すべての最新技術は、昔の Microsoft がすでに通った道の「リメイク版」だった**。

## 次に読む

- [ブラウザと VBA の双方向やり取り](/guides/two-way) — このツールで使える 5 つの選択肢と、IE 時代の方法
- [WebView2 登場秘話](/stories/DevHistory/webview2-story) — 純 VBA で WebView2 を埋め込むまで
- [BiDi 登場秘話](/stories/DevHistory/bidi-story) — このツールの BiDi も `Runtime.addBinding` で応答を受け取っている

## 参考資料

### IE・HTA の時代

* [A break from the past, part 2: Saying goodbye to ActiveX, VBScript, attachEvent…](https://blogs.windows.com/msedgedev/2015/05/06/a-break-from-the-past-part-2-saying-goodbye-to-activex-vbscript-attachevent/)（Microsoft Edge Blog、2015年5月。ActiveX と VBScript はどちらも 1996年に導入された、と説明している）
* [How to: Implement Two-Way Communication Between DHTML Code and Client Application Code](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2008/a0746166(v=vs.90))（Visual Studio 2008。`ObjectForScripting` / `window.external` と `Document.InvokeScript`）
* [Chapter 6 - Configuring the Active Desktop and Active Channels](https://learn.microsoft.com/en-us/previous-versions/cc768176(v=technet.10))（IE4 Resource Kit。Windows 98 のアクティブデスクトップ）
* [Introduction to HTML Applications (HTAs)](https://learn.microsoft.com/en-us/previous-versions/ms536496(v=vs.85))（HTA は完全に信頼されたアプリとして動き、ゾーンの設定に関係なく ActiveX を動かせる）
* [HTML Application](https://en.wikipedia.org/wiki/HTML_Application)（Wikipedia。1999年に IE5 とともに登場）

### 壁ができた時代

* [メールで広まる「I Love You」マクロウィルスに注意](https://internet.watch.impress.co.jp/www/article/2000/0508/love.htm)（INTERNET Watch、2000年5月）
* [Microsoft Security Advisory 2719662: Vulnerabilities in Gadgets Could Allow Remote Code Execution](https://technet.microsoft.com/library/security/2719662)（2012年7月。サイドバーとガジェットを無効にする Fix it）
* [Gadgets have been discontinued](https://web.archive.org/web/2015/http://windows.microsoft.com/en-us/windows/gadgets)（Microsoft のガジェットの配布ページ、Wayback Machine の写し）
* [A new approach to browser security: the Google Chrome Sandbox](https://blog.chromium.org/2008/10/new-approach-to-browser-security-google.html)（Chromium Blog、2008年10月）
* [Sandbox](https://chromium.googlesource.com/chromium/src/+/HEAD/docs/design/sandbox.md)（Chromium の設計文書）
* [ネイティブ メッセージング](https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging?hl=ja)（Chrome for Developers。拡張機能とローカルのアプリが標準入出力でやり取りする）

### IE の引退と組み込み IE

* [Internet Explorer 11 desktop app retirement FAQ](https://techcommunity.microsoft.com/blog/windows-itpro-blog/internet-explorer-11-desktop-app-retirement-faq/2366549)（2022年6月15日にサポート終了。MSHTML・WebOC・MSHTA は引き続きサポート）
* [「Edge」の「IE モード」は少なくとも2029年末までサポート ～Microsoftが改めて周知](https://forest.watch.impress.co.jp/docs/news/2145034.html)（窓の杜、2026年10月。MC1473151）
* [Microsoft、「Edge」の「IE モード」を制限 ～今夏よりセキュリティ攻撃の的に](https://forest.watch.impress.co.jp/docs/news/2054806.html)（窓の杜、2025年10月）
* [Deprecation of Legacy Web Browser Control](https://techcommunity.microsoft.com/blog/accessblog/deprecation-of-legacy-web-browser-control/4566944)（Access Blog、2026年11月。2027年5月18日、バージョン 2705 で廃止の予定）
* [Outlook Today and Folder Home Pages deprecation](https://support.microsoft.com/en-US/support/known-issues/outlook-today-and-folder-home-pages-deprecation)（Microsoft サポート。2026年8月に非推奨、2027年8月に無効）
* [Web connector troubleshooting](https://learn.microsoft.com/en-us/power-query/connectors/web/web-troubleshoot)（Power Query。`Web.Page` は IE ベース、`Web.BrowserContents` は WebView2 ベース）
* [OAuth with Microsoft Edge WebView2](https://learn.microsoft.com/en-us/power-query/oauth-edge-chromium)（Power Query。2021年10月から OAuth の画面が既定で WebView2）
* [サポート終了後もIEを使用している企業は12％](https://atmarkit.itmedia.co.jp/ait/articles/2208/02/news135.html)（@IT、2022年8月。バルテスの調査）
* [e-TaxがGoogle Chromeに一部対応、Chromiumベースの新しいMicrosoft Edgeでも利用可能に](https://forest.watch.impress.co.jp/docs/news/1254660.html)（窓の杜、2020年5月）
* [Windowsをご利用の方](https://www.jpki.go.jp/download/win.html)（公的個人認証サービス。利用者クライアントソフトの Edge / Chrome 版と拡張機能）

### WebView2

* [Announcing Microsoft Edge WebView2 General Availability](https://blogs.windows.com/msedgedev/2020/10/19/edge-webview2-general-availability/)（Microsoft Edge Blog、2020年10月19日）
* [interface ICoreWebView2](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2)（SDK 0.9.430 で登場。`AddHostObjectToScript`、`PostWebMessageAsJson`、`add_WebMessageReceived`）
* [interface ICoreWebView2SharedBuffer](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2sharedbuffer)（SDK 1.0.1661.34 で登場）
* [Shared Buffer Between Native Application Code and Script](https://github.com/MicrosoftEdge/WebView2Feedback/blob/4794d2fca7ab574f3546c9e05da45f4bf98c640b/specs/SharedBuffer.md)（SharedBuffer の仕様。背景に下の 2 つの要望を挙げている）
* [Communication via TypedArrays and ArrayBuffers（#89）](https://github.com/MicrosoftEdge/WebView2Feedback/issues/89)（2019年12月の要望。JSON と文字列では遅すぎる）
* [Binary messaging（#1005）](https://github.com/MicrosoftEdge/WebView2Feedback/issues/1005)（2021年2月の要望）

### Runtime.addBinding の経緯

* [Reland "[inspector] added Runtime.installBinding method"](https://chromium.googlesource.com/v8/v8/+/5cb11a17cdd1762f52ef5e46a5c963e33c5d5b06)（V8、Alexey Kozyatinskiy、2018年5月。`console.debug` の代わりの通信路として追加。差分のメソッド名は `addBinding`）
* [[inspector] add optional executionContextId arg to Runtime.addBinding](https://chromium.googlesource.com/v8/v8/+/e63f35bcaca4dcac5596d9707791a0c251ebea87)（V8、2018年6月）
* [fix(page): migrate exposeFunction from console.debug to Runtime.installBinding](https://github.com/puppeteer/puppeteer/commit/b20cde67c63a6344321c2fc40ef2d3cbf8c3ae24)（Puppeteer、2018年6月。[PR #2631](https://github.com/puppeteer/puppeteer/pull/2631)）
* [Puppeteer v1.0.0 の `lib/Page.js`](https://github.com/puppeteer/puppeteer/blob/v1.0.0/lib/Page.js)（`console.debug('driver:page-binding', ...)` と `Runtime.consoleAPICalled` による方式）
* [Puppeteer v1.6.0](https://github.com/puppeteer/puppeteer/releases/tag/v1.6.0)（Chromium 69.0.3477.0 を同梱）
* [chromium-bidi `src/bidiTab/bidiTab.ts`](https://github.com/GoogleChromeLabs/chromium-bidi/blob/bb98fab480403c61f322320f2ebd06c617ed0b90/src/bidiTab/bidiTab.ts)（`window.sendBidiResponse` は `Runtime.addBinding` で公開される）
* [セキュリティを強化するためのリモート デバッグ スイッチの変更](https://developer.chrome.com/blog/remote-debugging-port?hl=ja)（Chrome のブログ、Chrome 136。リモート デバッグを悪用した Cookie の抜き取りが増えた）
* [Let your Coding Agent debug your browser session with Chrome DevTools MCP](https://developer.chrome.com/blog/chrome-devtools-mcp-debug-your-browser-session?hl=ja)（Chrome のブログ、2025年12月、Chrome 144 以降）

### CDP の容量と、VBA の String

* [devtools_http_handler.cc](https://github.com/chromium/chromium/blob/81d56cf2fa29b4d5f4e2df10187a3a8c59ff9736/content/browser/devtools/devtools_http_handler.cc#L101-L102)（Chromium。WebSocket の送信 256MB・受信 100MB）
* [devtools_pipe_handler.cc](https://github.com/chromium/chromium/blob/e905c7b55b9cda9493b0c63a89192323ac942404/content/browser/devtools/devtools_pipe_handler.cc#L44-L45)（Chromium。パイプの受信 100MB、書き出しは 64KB ずつ）
* [http_connection.cc](https://github.com/chromium/chromium/blob/e8f2f6f1a988166465b719de8d3f803176cea827/net/server/http_connection.cc#L118-L121)（Chromium。上限を超えると「Too large write data is pending」として送れない）
* [String data type](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/string-data-type)（VBA。可変長の文字列は約 20 億（2^31）文字まで）

### Electron・Tauri

* [Electron のドキュメント](https://www.electronjs.org/docs/latest/)（Chromium と Node.js をバイナリに埋め込む）
* [Tauri の前提条件](https://v2.tauri.app/start/prerequisites/)（Windows では WebView2 で描画する）
