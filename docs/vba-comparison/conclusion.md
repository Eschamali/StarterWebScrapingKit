---
description: VBA 圏の3プロジェクトの用途別の使い分けと、同じ原典から出発した設計が何によって分かれたのかをまとめます。
---

# 使い分けと、分かれ道

## 観点別のまとめ

| 観点 | 優位 | 補足 |
| --- | --- | --- |
| イベント配送の仕組み | StarterWebScrapingKit | ネイティブ `RaiseEvent` × 4種、`sessionId` で自動分岐 |
| イベント処理の拡張性 | SWSK ≒ VBAChromeDevProtocol | 設計は互角。テンプレート化と型安全性で前者 |
| マルチタブ | StarterWebScrapingKit | `sessionId` 多重化、タブごとに独立オブジェクト |
| ポップアップ追跡の正確さ | VBAChromeDevProtocol | `openerId` で親子関係を厳密に辿る |
| 非同期実行 | StarterWebScrapingKit | 公開 API として `ExecuteCDPAsync` + `TakeResultCDP` |
| 生 CDP の叩きやすさ | VBAChromeDevProtocol | 236 ファイルの型付きドメイン層 |
| 要素操作 API の厚み | vba-cdp-webdriver ≒ SWSK | 前者は Selenium 作法、後者は待機・Shadow DOM が充実 |
| 単体タブでの手数の少なさ | vba-cdp-webdriver | 弱点が無関係になり、完成済みの便利機能が活きる |
| ドキュメント | StarterWebScrapingKit | 他2つは README とサンプルのみ |

## 用途で選ぶなら

**とりあえず何か自動化したい / 何をやるかまだ決まっていない**
StarterWebScrapingKit。マルチタブ・非同期・イベント拡張が全部「公式にサポート済み」の状態で揃っていて、詰まったときに読めるドキュメントがあります。

**単体タブで、フォーム入力とスクレイピングを手早く済ませたい**
vba-cdp-webdriver。マルチタブの弱点も拡張性の低さも、タブを1枚しか使わないなら発動しません。逆に `ClickAndThenAlertDialogErase`、`DownloadWatchStart`、`SetInterceptFileChooserDialog`、`EnableNetworkInterception` といった**よくあるイベント処理が完成品として載っている**ぶん、一番コードが短くなります。`WaitUntilVisible` / `FindElementByDeep` など Selenium 作法に馴染んだ人には API 名も直感的です。

**便利メソッドにない CDP コマンドを直接叩きたい**
VBAChromeDevProtocol。CDP 1.3 の全 45 ドメインが型付きクラスとして生成済みで、うち `Page` / `DOM` / `Network` / `Runtime` など主要 10 個は `cdp.Page.navigate` の形で即座に IntelliSense が効きます。残りも `New` して繋げば使えます。マニアックな要求に対しては3者で最強です。

**「ウィンドウを切り替える」という操作を素直に書きたい**
[SeleniumVBA](https://github.com/GCuser99/SeleniumVBA)。CDP 直叩きではなく本物の W3C WebDriver なので、ウィンドウハンドルの管理が仕様の標準機能として使えます。ただし `chromedriver.exe` が必要です。

## 3者を分けたもの

このコーナーで繰り返し出てきたのは、**設計のアイデアはほとんど同じ地点に立っていた**という事実でした。

- Pipe を本流に据え、後から WebSocket を足したという進化の順番が同じ
- `sessionId` による1本の接続の多重化という発想が同じ
- 非同期実行の下地（`nowait` / `ExecuteCDPAsync`）を内部に持っていた点まで同じ
- 「クリックで別タブが開く」問題への専用の仕組みを両方とも用意していた

VBAChromeDevProtocol に至っては、`registerEventHandler` の設計も、`openerId` による子ターゲット追跡も、CDP 公式仕様からの型安全なラッパー生成という**一番効くはずの投資**にまで着手していました。骨格は良い。

分かれたのは、その先です。

### 差が現れた場所

- 生成された `clsElement` は8メソッドのまま止まり、便利機能の層が育たなかった
- `nowait` という非同期の土台があるのに、公開 API に配線されないまま終わった
- ジェネレーターも README に「TODO: 荒削り」と書かれたまま更新が止まった
- `searchNull` の1文字ずつ回すループが、[原典のまま残り続けた](/vba-comparison/)

最後の1つが象徴的です。StarterWebScrapingKit 側の同じ関数には、こう書かれています。

```vb
' CDP messages received from chrome are null-terminated
' Updated: 25/10/25: Daniel Polak - new faster version
```

日付と名前が入っている。つまり**2025年10月に、誰かがこの関数を読んで、直した**ということです。

### 継続を支えたもの

`CDPBrowser.cls` の更新履歴コメントを見ると、その事情がもう少し見えてきます。

```
Updated: 22/03/23 Long Vh:      - Enhanced to dynamically retrieve installation path of the browser
         26/03/23 Long Vh:      - Made some quality-of-life changes
         25/05/25 Long Vh:      - set default profile to "CDP" so that it can work with the latest chrome update
         23/10/25 Daniel Polak: - Added missing space before --homepage
         22/01/26 Long Vh:      - Removed automatic reattachment from the start method as this restricts ...
         01/08/26 Eschamali     - iframeへの切り替えができるように改良
```

Pipe 通信の基盤を作った longvh211 氏本人が、2023年3月から2026年1月まで直接コミットし続けています。2025年10月からは Daniel Polak 氏も加わっている。同じ検索を VBAChromeDevProtocol のソース全体にかけても、こうした更新履歴コメント自体が1件も見つかりません。

これを「情熱の差」と呼ぶのは、少しもったいない気がします。もう少し構造的に言うなら。

1. **単独か、複数人か** ―― 一人開発はモチベーションが落ちた瞬間に止まりますが、複数人なら誰かが忙しくても別の誰かが繋ぎます
2. **ライブラリとして作るか、製品として作るか** ―― 「Example.xlsm を見てね」で止まったリファレンス実装より、使う人がいてフィードバックが返ってくる製品のほうが続きます
3. **ゼロから作るか、既にある良いものを組み合わせるか** ―― 「CDP 全ドメインを自動生成する」は正しい野心ですが、地味で終わりのない一人仕事でもあります

情熱そのものは有限で消耗します。差がついたのは**情熱を持続させる構造を用意できたかどうか**でした。

## そしてこの先に

同じ話は、もう一段スケールを上げても成立します。

```
同じ「CDP を method / id / sessionId で正しく捌く」というコアのアイデア
        ↓
    どれだけ継続的に投資されたか、という1本の軸だけが分岐点に
        ↓
VBAChromeDevProtocol ──→ StarterWebScrapingKit ──→ Playwright / Puppeteer
 （コアで力尽きた）      （少人数が投資し続けた）    （企業が何年も注ぎ込んだ）
```

テストの数、エラー分類の体系化、ブラウザ本体にパッチを当てるマルチエンジン対応 ―― [Puppeteer / Playwright との比較](/core-comparison/)で残った差も、「コアのアイデアが正しいかどうか」とは別の軸の話でした。

逆に言えば、**コアの設計思想さえ正しければ、個人・少人数のプロジェクトでも商用ツールと同じ土俵に立てる**。VBA という、ブラウザ自動化には普通誰も選ばない言語を選んでなお、です。

## 関連

- [概要と系譜](/vba-comparison/) — 2つの源流と、血縁の物的証拠
- [Puppeteer / Playwright とのコアロジック比較](/core-comparison/) — 同じ比較を Node 勢に対して
- [残る差分](/core-comparison/gaps) — 継続投資で埋まる差、埋まらない差
