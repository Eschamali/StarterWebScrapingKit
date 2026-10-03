---
description: Starter Web Scraping Kit の概要。Excel VBA 単体で CDP / WebDriver BiDi により Edge・Chrome を自動操作するキットの位置づけと特徴を紹介します。
---

# 概要

**Starter Web Scraping Kit** は、Excel/Access VBA だけで Chromium 系ブラウザ（Edge / Chrome）を自動操作するマクロブックです。  
VBAで、ChromeやEdgeを動かすと思い浮かぶのが、`SeleniumVBA`や`SeleniumBasic`です。しかし、これらは要管理者権限や `chromedriver.exe` 等の外部依存が必要になり、一種の壁となっています。

このツールは、徹底的に『外部依存ゼロ』にこだわり抜いて開発されました。

インストールも追加の参照設定も不要。リリースページからDLするだけで、Chrome と Edge を VBA から操作できます。  
特別な権限も、外部ツール(No Addin/PowerShell/driver.exe)の導入もいらない。「会社のPCでも、今日からそのまま動く」——それがこのツールのいちばんの価値です。  
――ええ、このドキュメントサイトをVitePressでビルドするために、私自身が`Node.js`をインストールする羽目になったという、最大の皮肉を噛み締めながらね🫥

このドキュメントサイトでは、**CDP** と **WebDriver BiDi** の使い方と API を扱います。

## できること

- 設定シートの内容でブラウザを起動し、ナビ・入力・クリック・JS 実行といった基本機能
- 既存セッションへの再接続（`reattach`）
- イベント処理
- `ExecuteCDP` / `ExecuteBiDi` による低レベル実行
- `--remote-debugging-pipe` , `--remote-debugging-port` , `WebView2`対応

## オブジェクトモデル

| Playwright(比較対象) | CDP | WebDriver BiDi |
| --- | --- | --- |
| Browser | [`CDPBrowser`](/api/cdp/CDPBrowser) | [`WebDriverBiDiMode`](/api/bidi/WebDriverBiDiMode) |
| Page | [`CDPContext`](/api/cdp/CDPContext) | [`WebDriverBiDiContext`](/api/bidi/WebDriverBiDiContext) |
| Locator / Element | [`CDPElement`](/api/cdp/CDPElement) | 当面は `jsEval` または [`UpgradeBiDiPlus`](/api/bidi/WebDriverBiDiContext#upgradebidiplus) |

入口は設定シート経由のワンライナーです。

::: code-group

```vb [CDP]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://kemono-friends.jp/"
t.ThisCDPBrowser.quit
```

```vb [BiDi]
Dim t As WebDriverBiDiContext
Set t = ShSetting01_StartBrowser.StartBiDiModeContext
t.navigate "https://kemono-friends-20170110.jp/"
t.ThisWebDriverBiDiMode.quit
```

:::

## IE の頃の書き方との対応

IE 自動操作の入門でよく見る「`Navigate` で指定 URL を開く」コードは、次のように書けます。

::: code-group

```vb [IE（かつての書き方）]
Dim objIE As InternetExplorer

'IE(InternetExplorer)のオブジェクトを作成する
Set objIE = CreateObject("InternetExplorer.Application")

'IE(InternetExplorer)を表示する
objIE.Visible = True

'指定したURLのページを表示する
objIE.Navigate "https://kemono-friends.jp/"

'完全にページが表示されるまで待機する
Do While objIE.Busy = True Or objIE.ReadyState <> 4
    DoEvents
Loop
```

```vb [Edge / Chrome（このツール）]
Dim t As CDPContext
Set t = ShSetting01_StartBrowser.StartCDPModeContext
t.navigate "https://kemono-friends.jp/"
t.ThisCDPBrowser.quit
```

:::

| IE | このツール |
| --- | --- |
| `CreateObject("InternetExplorer.Application")` + `Visible = True` | `StartCDPModeContext`（設定シートの内容でブラウザを起動） |
| `Navigate` + `Busy` / `ReadyState` の待機ループ | `navigate`（既定でページの読み込み完了まで待機するため、ループ不要） |
| `Quit` | `ThisCDPBrowser.quit` |

::: tip
待機方法は `navigate` の `WaitMode` 引数で変えられます（待たずに進める非同期指定も可能）。詳細は [ページ遷移](/guides/navigation) を参照してください。
:::

## どちらを使うか

迷ったら **CDP** から始めてください。要素操作（`CDPElement`）が揃っており、デモも豊富です。  
W3C BiDi 寄りに寄せたい、将来標準を先取りしたい場合は **BiDi**。足りない操作は `UpgradeBiDiPlus` や BiDi+（`goog:cdp.sendCommand`）で CDP に落とせます。  
詳細は [CDP と BiDi](/concepts/cdp-vs-bidi) を参照。  

## こんな方におすすめ

- IE終了後の代替手段を探している
- SeleniumVBA等が社内環境で使えなかった
- WebDriver管理をしたくない
- VBAだけでWeb自動化したい
- 社内システムへの入力作業を自動化したい
- スクレイピングや情報収集をしたい
- VBAを業務基盤とする実務担当者
- 新W3C準拠のWebDriverBiDiを触ってみたい

## 開発秘話

なぜ exe なしで WebDriver BiDi が動くのか——発見の経緯は [BiDi 登場秘話](/stories/bidi-story) にまとめています。

## UserForm にモダンブラウザを載せる

IE コントロールの代替として、Edge / WebView2 を UserForm に載せる手法は [UserForm コーナー](/userform/intro) へ。

## 次のステップ

1. [はじめに](/getting-started) — 保護ビュー解除と Hello World
2. [アーキテクチャ](/concepts/architecture) — クラスの役割
3. [設計思想](/concepts/design-philosophy) — povo 2.0 スタイル
4. [ページ遷移](/guides/navigation) — 最初のガイド
