# なぜ `WebDriverBiDi.exe` なしでブラウザ自動化ができるとわかったのか

::: tip 登場秘話
〜公式ドライバーの闇を暴いた、土日の記録〜
:::

## すべては、ある海外の議論から

![海外フォーラムでの議論](/img/story/1.png)

すべては、ある海外の議論で見かけたこの意見から始まりました。

> 「CDP（Chrome DevTools Protocol）はChrome専用の独自仕様。将来性は WebDriver BiDi が上だ。」

言っていることは正しい。でも、現場のリアルは少し違います。企業環境において自動化を阻む本当の壁は「ブラウザのインストール禁止」よりも、**「プリインストール以外のexeやNode.jsを情シスが許可しない」** という壁でした。

Windowsの標準搭載とExcelのインフラ化が後押しし、以下の組み合わせで環境はすでに整っていました。

- REST WebAPI → `WinHTTP 5.1`
- ブラウザ自動操作 → `Edge-CDP via Pipe`
- WebSocket通信 → `Winhttp.dll`

それでも……**「将来性はBiDiが上」** という言葉が頭から離れなかった。

---

## 壁 — `msedgedriver.exe` という巨大な存在

![msedgedriverの壁](/img/story/2.png)

ネットを調べるたびに、同じ言葉が出てくる。

> 「まず、対応するWebDriver — Edgeなら `msedgedriver.exe` をダウンロードします」

ExcelでBiDiを使う記事があっても、結局は *exe × WebSocket* の構成に依存してしまう。  
「やっぱりexeが要る。情シスの壁は越えられない……」と一度は諦めかけた。

それでも手が止まらず調査を続けた数日後、あるリポジトリに辿り着いた。

![核となるリポジトリ](/img/story/3.png)

🔗 [GoogleChromeLabs / chromium-bidi](https://github.com/GoogleChromeLabs/chromium-bidi)

いかにもChromium公式のBiDiリポジトリ。しかしREADMEには `Node.js`、`npm`…「結局これも外部依存か」と落胆した。

---

## 転機 — AIが明かした真実。主役は「exe」ではなかった

![Google AI Pro 特典メール](/img/story/4.png)

そんな中、「**Google AI Pro を3か月お試し**」というメールが届いた。早速有効化し、[Antigravity](https://antigravity.google/) といったAIにコードを読み込ませて解説してもらう機能を発見した。

![Antigravityのホームページ](/img/story/5.png)

::: tip 閃き
**「AIに chromium-bidi のソースを読ませて、Excelで完全再現できないか？」**
:::

そして、衝撃の事実が判明した。`Node.js` や `.exe` は単なる「**運び屋（橋渡し役）**」に過ぎなかった。AIは言った——

> 「`mapperTab.js` という巨大なJavaScriptファイルこそが、BiDiの心臓部です。」

![BiDiの心臓部ソースコード](/img/story/6.png)

---

## 手順 — 実現のための「5ステップ」

![大まかな手順](/img/story/7.png)

1. **JSの入手:** `npm` や `JSDelivr` などのCDNから `mapperTab.js` を取得する。
2. **特権の付与:** CDPコマンド `Target.exposeDevToolsProtocol` でタブにブラウザ操作の特権を与える。
3. **窓口の確保:** `Runtime.addBinding` でVBAと通信するための受け取り口を確保。
4. **注入と起動:** `mapperTab.js` をタブに注入し、`Runtime.evaluate` でBiDiを起動する。
5. **非同期通信:** `Runtime.bindingCalled` イベントをキャプチャし、BiDiの非同期レスポンスを受け取る。

もうお分かりだろう。**「mapperTab.js というブラウザ内で動くプログラムが、BiDiコマンドをCDPに翻訳する作業をぜ〜〜んぶ肩代わりしていたのだ。」**

---

## 封印 — Excelのセルにブラウザの心臓部を閉じ込める

主役はバイナリ（exe）ではなく、テキストデータ（js）だった。  
**そうです。テキストなら、Excelのセルに置けちゃうのです。**

![Excelのセルにブラウザの心臓部を封印](/img/story/8.png)

数万行のスクリプトでも複数セルに分割して格納し、VBAの `Join` 関数で結合してブラウザへ注入できる。バイナリ（exe）は情シスに即ブロックされるが、テキストデータなら**「ただのExcelファイル」**としてパスできるのだ。

::: tip 到達点
**ついに、Excel単体でBiDiコマンドが実行できるようになりました！**
:::

土日を完全に溶かし、BiDi版の低レベル制御機能（`ExecuteBiDi`, `ExecuteBiDiAsync`, `TakeEvents`）を `WebDriverBiDiMode` / `WebDriverBiDiContext` として作り上げた。

![完成](/img/story/12.png)

さらに、SeleniumVBAにあった「自動更新機能」も独自実装。`jsdelivr.com` のAPIを叩くことで:

![自動更新のサイト](/img/story/9.png)

- 最新バージョンのチェック
- `mapperTab.js` 自体の自動ダウンロード

がVBA単体で完結。SeleniumVBAが「フォルダに `webdriver.exe` を配置」するのに対し、このツールは **「ExcelのテーブルにJSのテキストを上書き」** するだけ。情シスの監視をすり抜ける*究極のステルス仕様*だ。

---

## 疑惑 — 公式 `msedgedriver.exe` も同じハックをしているのか？

夢は叶った。しかし一つの怖い疑問が湧いた。

> 「本当に公式の msedgedriver.exe も、ただの橋渡し役なのか？」

もしexeがもっと高度なネイティブロジックで動いていたら、自分の作ったものは「非公式の迂回ルート」になってしまう。

そこで、[SeleniumVBAのBiDi拡張版](https://github.com/hanamichi77777/WebDriverBiDi-via-VBA-test)を使い、**公式ドライバーが裏で何をやっているか、この目で確認することにした。**

![SeleniumVBAにWebDriver BiDi機能を付けた拡張版](/img/story/10.png)

起動してみたが、画面に「BiDi-CDP Mapper is controlling this tab」というタブは見当たらない。  
「やっぱり違う戦法か」と絶望しかけたが、AIからヒントが届いた。

![AIからの新たな手がかり](/img/story/11.png)

> 「新しいタブを **非表示（type: other）** として生成している可能性があります。`edge://inspect/#devices` で確認できるはずです」

早速ブラウザのデバッグ画面を開くと……

![デバッグ画面を開くと](/img/story/13.png)

1つしか開いてないのにターゲットが3つ？ クリックしてみると……**あるじゃありませんか！**

![あるじゃありませんか](/img/story/14.png)

非表示タブなので画面は描画されないが、`outerHTML`をコピーしてファイル化してみると、BiDiコマンドが処理されている痕跡が残っていた。

| outerHTML をコピー | ファイル化 | BiDi 処理の痕跡 |
| --- | --- | --- |
| ![outerHTMLをコピー](/img/story/15.png) | ![ファイル化してみると](/img/story/16.png) | ![BiDiコマンドを処理している痕跡](/img/story/17.png) |

さらにAIからの助言が続いた。**「exe をテキストエディタで強引に開けば証拠が見つかるかもしれません」**

「どうかプレーンテキストで残ってますように……！」と祈りながら、`msedgedriver.exe` をバイナリエディタで強引に開き、検索をかけた。

::: tip 発見
**ヒットした！**
:::

著作権表示と共に `<!DOCTYPE html><title>BiDi-CDP Mapper</title>...` という生々しいコードが、数万行にわたってハードコードされていた！

![バイナリエディタで検索結果](/img/story/18.png)

あの重厚な公式ドライバーも、裏では私がVBAでやったのと全く同じ **「JSの翻訳機を隠しタブに注入する」** という泥臭いハックをやっていた。「この土日の作業は無駄ではなかった」と心から納得できた。

---

## 結論 — そして、未来へ

WebDriver BiDi 自体はまだβ版。将来的には `mapperTab.js` が不要になる日も来るかもしれない。しかし、AIをフル活用してこの「ロマンティックなツール」を作り上げたことは、最高の勉強になった。

言語も環境（バイナリ、Node.js、Excelマクロ）も全く違うのに、**ブラウザの裏口（CDP）を開けて翻訳機を忍ばせる**という本質的なアプローチは、見事に共通していた。

| 方式 | 仕組み |
| --- | --- |
| exe によるオートメーション | `msedgedriver.exe` などが内部C++の文字列として隠し持ち、起動時に注入する |
| Node.js によるオートメーション | Google Chrome Labs の chromium-bidi リポジトリから直接呼び出される |
| VBA によるオートメーション | Excelのセルにテキストとして封印され、VBAマクロから直接ブラウザへ送り込まれる |

最新のWeb標準規格の裏側を支えているのが、たったひとつの巨大なJavaScriptファイル。それが「Excel VBA」というレガシーな環境にもピタリと当てはまる。

::: tip フィナーレ
技術の最先端と普遍性を同時に味わえた、最高にエキサイティングで、**ハッカーとしてのロマンに溢れた週末だった！**
:::

## 次に読む

- [アーキテクチャ](/concepts/architecture) — mapper が載る層の位置づけ
- [CDP と BiDi](/concepts/cdp-vs-bidi) — どちらを使うか
- [生プロトコル拡張](/guides/extend-raw-protocol) — `ExecuteBiDi` / BiDi+
