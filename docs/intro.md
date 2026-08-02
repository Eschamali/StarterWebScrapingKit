---
description: Starter Web Scraping Kit の概要。Excel VBA 単体で CDP / WebDriver BiDi により Edge・Chrome を自動操作するキットの位置づけと特徴を紹介します。
---

# 概要

**Starter Web Scraping Kit** は、Excel VBA だけで Chromium 系ブラウザ（Edge / Chrome）を自動操作するマクロブックです。Selenium の `chromedriver.exe` は不要です。  
徹底的に『外部依存ゼロ』にこだわり抜いて開発されました。  
――ええ、このドキュメントサイトをVitePressでビルドするために、私自身が`Node.js`をインストールする羽目になったという、最大の皮肉を噛み締めながらね🫥

このドキュメントサイトでは、**CDP** と **WebDriver BiDi** の使い方と API を扱います。

## できること

- 設定シートの内容でブラウザを起動し、ナビ・入力・クリック・JS 実行といった基本機能
- 既存セッションへの再接続（`reattach`）
- イベント処理
- `ExecuteCDP` / `ExecuteBiDi` による低レベル実行
- `--remote-debugging-pipe` , `--remote-debugging-port` 両対応

## オブジェクトモデル

| Playwright(比較対象) | CDP | WebDriver BiDi |
| --- | --- | --- |
| Browser | [`CDPBrowser`](/api/cdp/CDPBrowser) | [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) |
| Page | [`CDPContext`](/api/cdp/CDPContext) | [`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) |
| Locator / Element | [`CDPElement`](/api/cdp/CDPElement) | 当面は `jsEval` または [`ConvertToCDPContext`](/api/bidi/WebDriverBiDiContext#converttocdpcontext) |

入口は設定シート経由のワンライナーです。

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://kemono-friends.jp/"
t.InheritanceCDPBrowser.quit
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext
t.navigate "https://kemono-friends-20170110.jp/"
t.InheritanceWebDriverBiDiMode.quit
```

:::

## どちらを使うか

迷ったら **CDP** から始めてください。要素操作（`CDPElement`）が揃っており、デモも豊富です。

W3C BiDi 寄りに寄せたい、将来標準を先取りしたい場合は **BiDi**。足りない操作は `ConvertToCDPContext` や BiDi+（`goog:cdp.sendCommand`）で CDP に落とせます。

詳細は [CDP と BiDi](/concepts/cdp-vs-bidi) を参照。

## 開発秘話

なぜ exe なしで WebDriver BiDi が動くのか——発見の経緯は [BiDi 登場秘話](/stories/bidi-story) にまとめています。

## UserForm にモダンブラウザを載せる

IE コントロールの代替として、Edge / WebView2 を UserForm に載せる手法は [UserForm コーナー](/userform/intro) へ。

## 次のステップ

1. [はじめに](/getting-started) — 保護ビュー解除と Hello World
2. [アーキテクチャ](/concepts/architecture) — クラスの役割
3. [設計思想](/concepts/design-philosophy) — povo 2.0 スタイル
4. [ページ遷移](/guides/navigation) — 最初のガイド
