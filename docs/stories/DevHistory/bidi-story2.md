---
description: 「BiDi 登場秘話」のその後。Chromium の WebDriver BiDi まわりで見つかった動き（Mapper の同梱、タブ以外で動かす変更、BiDi の WebSocket の試作など）を、日付順に追記していく更新記録です。
---

# BiDi 登場秘話・その後

<PromoMovie src="/promo/ReiwaIE_BiDiStory2.html" title="令和のIE 技術秘話 ― BiDi 登場秘話・その後" />

[BiDi 登場秘話](/stories/DevHistory/bidi-story)は、2026年10月の執筆時点のお話です。

その後、Chromium の WebDriver BiDi まわりが、思った以上に動き始めました。本編は執筆時点の内容のまま残し、このページに、見つかった動きを日付順に足していきます。新しい記録は、下に追記していきます。

::: info このページの読み方
* 日付は、筆者が確かめた日です。変更が作られた日・取り込まれた日は、それぞれの記録の中に書いています
* 「レビュー中」「実験用」の変更は、そのまま製品に入るとは限りません。状態が変わったら、新しい日付の記録で追いかけます
* 推測を書くときは、推測だと分かるように書いています
:::

## 2026年11月17日：Mapper が Chrome 本体に同梱された

### きっかけ

`chrome://inspect/#remote-debugging` まわりの変更を追いかけていたら、同じファイル（`devtools_http_handler.cc`）を触る変更の一覧に、こんなものが並んでいました。

| 日付 | 変更 | 中身 |
| --- | --- | --- |
| 2026-11-11 | [8786163](https://chromium-review.googlesource.com/c/chromium/src/+/8786163)「[bidi] Bundle ChromeDriver BiDi mapper into content resources」 | ビルドの設定 `bundle_chrome_bidi_mapper` を追加（デスクトップでは有効）。`mapperTab.js` を縮めて Brotli で圧縮し、Chrome 本体の `content_resources.pak` に `IDR_CHROME_BIDI_MAPPER_JS` として同梱する |
| 2026-11-12 | [8793639](https://chromium-review.googlesource.com/c/chromium/src/+/8793639)「[DevTools] Add /bidi-mapper.js endpoint to serve bundled BiDi mapper」 | デバッグ用の HTTP サーバーが、`/bidi-mapper.js` でその同梱物を返す。設定が無効なビルドでは 404 を返す |

本編の「`BiDi-CDP Mapper`(mapperTab.js)の未来予想」で、「Chrome 本体に Mapper を持たせる」案は「ダメなんです」と書きました。それが、形を変えて始まっていたのです。

### 2022年の設計文書の難点と並べると

本編で紹介した2022年の設計文書「CDP-BiDi Mapper code lifecycle」は、Chrome に Mapper を持たせる案の難点を3つ挙げていました。今回の変更と並べると、こうなります。

* **バイナリサイズ**：変更の説明によると、縮めて圧縮した Mapper で増えるのは約100KB（`resources.pak` の0.4%）です。しかも、同梱するのはデスクトップ版だけです
* **CDP のコマンドとして承認を得る難しさ**：今回の形は CDP のコマンドではなく、デバッグ用の HTTP サーバーの URL です
* **バージョンを合わせる問題**：ブラウザのビルドに同梱されるので、そのブラウザと同じビルドの Mapper が手に入ります

### Chrome Canary で試してみた

こうなると試したくなるのが人情です。2026年11月17日、Chrome Canary と Edge Canary を入れて、それぞれのデバッグ用ポートに `GET /bidi-mapper.js` を送ってみました。

| ブラウザ | バージョン | 結果 |
| --- | --- | --- |
| Chrome Canary | 160.0.8166.2 | **200**。685,153バイトの `text/javascript` |
| Edge Canary | 160.0.4408.0 | **404** |

Chrome Canary が返したのは、縮めた（minify した）`mapperTab.js` でした。npm の `mapperTab.js`（160.0.8166-0 で 1,033,272バイト）より3割ほど小さくなっています。Google と Microsoft の著作権表示（`Copyright 2021 Google LLC.`、`Copyright (c) Microsoft Corporation.`）も残っていました。中にある BiDi のコマンド名・イベント名（`browsingContext.navigate` など）を抜き出して npm の 160.0.8166-0 と比べると、どちらも116件で、違いはありませんでした。

ちなみに、npm には 160.0.8166.2 にぴったり合う `160.0.8166-2` がありませんでした。npm から取るなら近い版を選ぶしかありませんが、ブラウザ自身から取れば、そのブラウザと同じビルドの Mapper がそのまま手に入ります。

Edge Canary の404は、Edge のビルドではこの設定が無効になっているのか、Edge Canary の元になった Chromium がまだこの変更より古いのか、今のところ区別できません。

### 変わっていないこと

* **Mapper が動く場所は、今もタブです。** 今回変わったのは「Mapper がどこに入っているか」で、「どこで動くか」ではありません。取ってきた Mapper をタブに注入して動かす流れは、これまでと同じです
* **取れるのは、デバッグ用の HTTP サーバーがあるときだけです。** `--remote-debugging-pipe` で起動した場合は HTTP サーバーが無いので、`/bidi-mapper.js` は使えません
* **ChromeDriver は、まだ同梱物を使っていません。** 変更の題名には「ChromeDriver BiDi mapper」とありますが、ChromeDriver の最新のソース（2026年11月17日時点の main）では、Mapper は今も ChromeDriver 自身に埋め込んだ定数 `kMapperScript` から取っています（[session\_commands.cc](https://chromium.googlesource.com/chromium/src/+/a73da91b4016bcc4d25ed67ecf86cdf09a42f1a2/chrome/test/chromedriver/session_commands.cc#441)）。起動引数 `--bidi-mapper-path` でファイルを渡したときだけ、そちらを読みます。`kMapperScript` の元は、`third_party/chromium-bidi` から作った `mapperTab.js` を縮めて C++ の定数にしたものです（[BUILD.gn](https://chromium.googlesource.com/chromium/src/+/a73da91b4016bcc4d25ed67ecf86cdf09a42f1a2/chrome/test/chromedriver/BUILD.gn#79) の `minify_bidimapper` と `embed_bidimapper_in_cpp`）。ChromeDriver のソースには、`/bidi-mapper.js` や同梱物（`IDR_CHROME_BIDI_MAPPER_JS`）を参照している箇所はありませんでした

### タブではない場所で Mapper を動かす変更（レビュー中）

同梱物を使う側の動きも、調べてみるとすでに始まっていました。

[8817570](https://chromium-review.googlesource.com/c/chromium/src/+/8817570)「[bidi] Add BiDi mapper host for the chrome.driver extension API」（2026年11月13日作成、11月17日時点でレビュー中）は、`content::BidiMapperHost` という仕組みを追加します。変更の説明によると、次のような内容です。

* 同梱の Mapper を、**非表示の `chrome-untrusted://bidi-mapper` というページ**（特権付きの WebContents）で動かし、BiDi のメッセージを中継する
* Mapper のページに置いたメッセージのフィルターがセッションを管理し、`goog:cdp`、`script.evaluate`、`script.callFunction` を拒否する
* 起動時の引数で `ChromeDriverExtensionApi` を有効にしたときだけ動く

名前のとおり、Chrome 拡張機能の API（`chrome.driver`）向けの仕組みです。それでも、本編の結論で「最適な場所が存在しない」と書いた「タブ以外のどこか」の候補が、具体的な形で出てきたことになります。本編で検討した Service Worker（止まってしまう問題）でも、普通のタブでもない、第3の場所です。

### 計画の中身：拡張機能向けの `chrome.driver`

8817570 のきっかけになったバグ [585298118](https://issues.chromium.org/issues/585298118)「Expose the bundled BiDi mapper to extensions via a chrome.driver API」（2026年11月13日作成、P1、作業中）には、計画の中身が書かれていました。担当者は、同梱の変更（8786163・8793639）と同じ人です。

狙いは、次の一文です。

> Expose it to extensions so that an extension can automate the browser over WebDriver BiDi without ChromeDriver or a remote debugging port.

拡張機能が、**ChromeDriver もデバッグ用のポートも使わずに**、WebDriver BiDi でブラウザを自動化できるようにする、という意味です。

API の案（バグの言葉では「Initial API design scetch」）は、次のとおりです。

* `chrome.driver.connect()` が、`onMessage`・`sendMessage(message)`・`close()` を持つ入口を返す。やり取りするのは生の BiDi の JSON
* 接続ごとに、非表示の WebContents で Mapper を1つずつ動かす。1つの接続が1つの BiDi のセッションを受け持ち、そのセッションからはブラウザ全体が見える
* 拡張機能には、新しい「driver」の権限が要る

制約も書かれています。

* `ChromeDriverExtensionApi` という機能のフラグで切り替える。**既定はオフで、起動時の引数でしか有効にできず、安全でないフラグの扱い**
* Mapper は手を加えずに使う（ChromeDriver と共有するため）
* 生の CDP を使う口（`goog:cdp`）は拒否する
* 今のところ、デスクトップ版だけ

なお、バグの案で拒否すると書かれているのは `goog:cdp` だけです。実際の変更（8817570）の説明では、`script.evaluate` と `script.callFunction` も拒否すると書かれています。作っていく途中で、範囲を絞ったのかもしれません（ここは筆者の推測です）。

### `chrome.debugger` の BiDi 版

拡張機能から CDP を使う API は、前から `chrome.debugger` がありました。`chrome.driver` は、その BiDi 版にあたります。

| | `chrome.debugger` | `chrome.driver`（計画中） |
| --- | --- | --- |
| 話せるプロトコル | CDP | WebDriver BiDi |
| 拡張機能に要る権限 | `debugger` | 新しい `driver` |
| 状態 | 正式な API | フラグでのみ有効（既定でオフ、安全でないフラグ扱い） |

後で紹介する実験用の 8634745 では、BiDi を `chrome.debugger` の `targetId: "bidi"` で使う形でした。585298118 では、それを**別の API と別の権限**に分ける設計になっています。

ただし、「Chrome に BiDi が組み込まれた」と言い切るには、1つ注意が要ります。

* **API としては、Chrome に組み込みです。** 拡張機能から見れば、Chrome が最初から BiDi の入口を持っている形になります
* **中身は、今と同じ Mapper（JavaScript）です。** 非表示のページで Mapper を動かし、BiDi を CDP に翻訳しています

Firefox のようにブラウザが BiDi をそのまま話すのではなく、「Chrome が、翻訳係（Mapper）を自分の中で雇って、BiDi の窓口を開く」形です。窓口から見ればネイティブ、中身は翻訳、という二重構造です。

ちなみに、`chrome.debugger` を使うと、ブラウザの上部に「〇〇がこのブラウザのデバッグを開始しました」というバーが出ます。`chrome.driver` で同じような表示が出るかどうかは、バグにも変更の説明にも書かれていませんでした。

### Native Messaging と組み合わせると

`connect()` が `onMessage`・`sendMessage`・`close` を持つ入口を返す形は、Native Messaging の `chrome.runtime.connectNative()` が返す `Port`（`onMessage`・`postMessage`・`disconnect`）によく似ています。

| | Native Messaging | `chrome.driver`（計画中） |
| --- | --- | --- |
| つなぐ相手 | 拡張機能 ⇔ **ブラウザの外のアプリ**（標準入出力） | 拡張機能 ⇔ **ブラウザの中の Mapper**（BiDi） |
| 入口 | `chrome.runtime.connectNative()` | `chrome.driver.connect()` |

役割は別物ですが、2つをつなぐと、外のアプリから BiDi でブラウザを動かす道ができます。

```text
外のアプリ（たとえば VBA）
   │  標準入出力（Native Messaging）
   ▼
拡張機能
   │  chrome.driver.connect()（生の BiDi の JSON）
   ▼
非表示のページで動く Mapper
   │  CDP
   ▼
ブラウザ全体
```

デバッグ用のポートを開かないので、`chrome://inspect/#remote-debugging` の承認ダイアログも出ません。拡張機能と Native Messaging で承認ダイアログを避ける方法は、これまでも `chrome.debugger`（CDP）で使われてきました。その先が BiDi になる形です。

ただ、今の計画では「安全でないフラグ」を起動時に付ける前提なので、普段使いのブラウザにそのまま使える状態ではありません。当面は実験向けです。

### Chrome に BiDi の WebSocket の入口を持たせる試作（取り込まない前提）

もう1つ、[8634745](https://chromium-review.googlesource.com/c/chromium/src/+/8634745)「[chromium-bidi] Bundle ChromeDriver BiDi mapper with Chrome」（2026年10月23日作成）という変更もありました。説明の冒頭には「EXPERIMENTAL DO NOT LAND!」（実験用。取り込まない）と書かれています。

* Mapper を Chrome に同梱し、デバッグ用の HTTP サーバーに **`/session`、`/session/<id>`、`/devtools/bidi` という BiDi の WebSocket の入口**を作る
* Chrome 拡張機能の `chrome.debugger` API からも、`targetId: "bidi"` で BiDi を使えるようにする
* Mapper は非表示の特権付き WebContents で動かし、`session.new`・`session.end` などの BiDi のセッションの流れを Chrome 側で受け持つ

つまり、**Firefox のように、Chrome に WebSocket で直接 BiDi を話させる試作**です。本編の「将来、ChromiumはWebDriverBiDiをネイティブに扱えるのか？」では「示唆は見当たらなかった」と書きましたが、少なくとも試作は存在していました。

取り込まない前提の実験なので、このまま製品になるわけではありません。ただ、11月に取り込まれた同梱（8786163）や `/bidi-mapper.js`（8793639）、レビュー中の 8817570 は、この試作の一部を切り出したもののようにも見えます（ここは筆者の推測です）。

### 「なんでもOKの BiDi」と「安全モードの BiDi」

8817570 と 8634745 を並べると、Chrome 側の考え方が見えてきます。8634745 の説明には、次の一文があります。

> blocking `goog:cdp.*` commands and events for untrusted clients

「信頼できないクライアントには `goog:cdp.*` を使わせない」という意味です。接続してきた相手を信頼できるかどうかで、使える BiDi の範囲を変える考え方のようです。

| | なんでもOKの BiDi | 安全モードの BiDi |
| --- | --- | --- |
| 使う相手 | ChromeDriver や、このツールのように、CDP でつないで自分で Mapper を注入するクライアント | 拡張機能（`chrome.driver` など）や、信頼できないクライアント |
| `goog:cdp.*`（BiDi の中から CDP を使う口） | 使える | **使えない** |
| `script.evaluate`・`script.callFunction` | 使える | **使えない**（8817570） |
| そのほか | ― | 接続した相手が触ってよいタブや URL の範囲を守らせる（8634745） |

面白いのは、`script.evaluate` と `script.callFunction` が、**WebDriver BiDi の標準のコマンド**だという点です。安全モードは、「標準の BiDi から、危ない部分を削った一部分」になります。標準の仕様そのものには、こういう段階の考え方はありません。Chrome 独自の線引きです。

このツールは、`WebDriverBiDiContext.UpgradeBiDiPlus` で `CDP○○.cls` を使っている時点で、はっきり「なんでもOK」側です。CDP で直接つないで自分で Mapper を注入しているので、Chrome 側に安全モードの仕組みが入っても、すぐに影響を受ける立場ではありません。安全モードがいずれ「デバッグ用のポートにつないだ外部のクライアント」にも広がるかどうかは、まだ分かりません。

### ブラウザにつなぐルートの整理

ここまでの話を、ブラウザにつなぐルートごとに並べるとこうなります。

| ルート | 入口 | 話せるもの | 状態 |
| --- | --- | --- | --- |
| 1. 起動時の引数 | `--remote-debugging-pipe` / `--remote-debugging-port` で、新しくブラウザを起動する | CDP（BiDi は自分で Mapper を注入する） | 前から |
| 2. 承認モード | `chrome://inspect/#remote-debugging` で、普段使いのブラウザのポートを開く。接続のたびに承認ダイアログが出る | CDP（BiDi は自分で Mapper を注入する） | 2025年11月から |
| 3. 拡張機能 | `chrome.debugger` | CDP | 前から |
| 　 | `chrome.driver` | **BiDi** | **工事中**（フラグでのみ有効） |

取り込まない前提の試作（8634745）では、ルート1・2のデバッグ用 HTTP サーバーにも BiDi の WebSocket の入口を作っていました。

Chrome 側の工事をまとめると、次の3つになります。

* **Mapper を Chrome 本体に持たせる**（済み：同梱と `/bidi-mapper.js`）
* **Mapper をタブ以外の場所で動かす**（工事中：非表示の `chrome-untrusted://bidi-mapper`）
* **BiDi の入口をルートごとに作る**（拡張機能は工事中、WebSocket は試作だけ）

今までは「BiDi を使いたければ、ChromeDriver を使うか、自分で Mapper を注入する」しかありませんでした。それを、「Chrome 自身が BiDi の入口を持つ」方向へ少しずつ寄せていると読めます。

### 本編の予想と比べて

本編の「未来予想」の結論では、「いつかタブではないどこかで`BiDi-CDP Mapper`が動く」と書きました。それより先に「Mapper を Chrome 本体が配る」が現実になり、「タブではない場所で動かす」もレビュー中です。

「将来、ChromiumはWebDriverBiDiをネイティブに扱えるのか？」では、「一生、ネイティブ搭載はしない」と予想しました。ところが、同梱の変更のきっかけになったバグ [584772915](https://issues.chromium.org/issues/584772915)「Bundle ChromeDriver BiDi mapper into content resources」（2026年11月12日作成、修正済み）の説明には、次の一文がありました。

> For Chrome to support WebDriver BiDi directly, the minified, Brotli-compressed BiDi mapper also needs to be bundled into Chrome's content resources.

「Chrome が WebDriver BiDi を直接サポートするには、縮めて Brotli で圧縮した Mapper を、Chrome の content resources にも同梱する必要がある」という意味です。一連の工事が「Chrome が BiDi を直接サポートする」ことを目指しているのは、Google 側の言葉として残っています。

ただ、中身は今も Mapper による翻訳です。拡張機能の API は安全でないフラグの奥にあり、BiDi の WebSocket の入口は取り込まない前提の試作です。予想が外れたのかどうかは、まだ判断できません。

Mapper の置き場所を巡る話は、まだまだ続きそうです。

## 参考資料

ソースコードへのリンクは、確かめた時点の版に固定しています。

* Chromium の変更: [8786163](https://chromium-review.googlesource.com/c/chromium/src/+/8786163)「[bidi] Bundle ChromeDriver BiDi mapper into content resources」（2026-11-11）
* Chromium の変更: [8793639](https://chromium-review.googlesource.com/c/chromium/src/+/8793639)「[DevTools] Add /bidi-mapper.js endpoint to serve bundled BiDi mapper」（2026-11-12）
* Chromium の変更（レビュー中）: [8817570](https://chromium-review.googlesource.com/c/chromium/src/+/8817570)「[bidi] Add BiDi mapper host for the chrome.driver extension API」（2026-11-13 作成）
* Chromium の変更（実験用、取り込まない前提）: [8634745](https://chromium-review.googlesource.com/c/chromium/src/+/8634745)「[chromium-bidi] Bundle ChromeDriver BiDi mapper with Chrome」（2026-10-23 作成）
* Chromium の Issue: [584772915](https://issues.chromium.org/issues/584772915)「Bundle ChromeDriver BiDi mapper into content resources」（2026-11-12 作成、修正済み）
* Chromium の Issue: [585298118](https://issues.chromium.org/issues/585298118)「Expose the bundled BiDi mapper to extensions via a chrome.driver API」（2026-11-13 作成、作業中）
* Chromium（2026-11-17 時点の main）: [chrome/test/chromedriver/session\_commands.cc](https://chromium.googlesource.com/chromium/src/+/a73da91b4016bcc4d25ed67ecf86cdf09a42f1a2/chrome/test/chromedriver/session_commands.cc#441)（`kMapperScript` と `--bidi-mapper-path`）、[chrome/test/chromedriver/BUILD.gn](https://chromium.googlesource.com/chromium/src/+/a73da91b4016bcc4d25ed67ecf86cdf09a42f1a2/chrome/test/chromedriver/BUILD.gn#79)（`minify_bidimapper`・`embed_bidimapper_in_cpp`）
* Chrome for Testing: [last-known-good-versions.json](https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json)（各チャンネルのバージョンとリビジョン）
* 設計文書: [CDP-BiDi Mapper code lifecycle](https://docs.google.com/document/d/1_UTBlOa5GCdf075ni4hDVgBSHkgeblfP_g_hwxqRhjg)（Maksim Sadym、2022-01-03）

## 次に読む

- [BiDi 登場秘話](/stories/DevHistory/bidi-story) — 本編。2026年10月の執筆時点のお話
- [アーキテクチャ](/concepts/architecture) — mapper が載る層の位置づけ
- [CDP と BiDi](/concepts/cdp-vs-bidi) — どちらを使うか
