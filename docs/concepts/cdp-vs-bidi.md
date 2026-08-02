---
description: Chrome DevTools Protocol（CDP）と WebDriver BiDi の違い・選び方・併用パターンを早見表付きで比較するガイドです。
---

# CDP と BiDi

同じキットに **2 つのプロトコル面** があります。用途で選び、必要なら混ぜます。

## 早見表

| | CDP | WebDriver BiDi |
| --- | --- | --- |
| 入口 | `StartCDPModeContext` | `StartBiDiModeContext` |
| ページ型 | `CDPContext` | `WebDriverBiDiContext` |
| 要素型 | `CDPElement`（充実） | 高レベル要素 API は限定。`jsEval` か CDP 変換 |
| イベント | `BrowserEvents` + `Network.enable` 等 | `sessionSubscribe` + `BiDiEvents` |
| 生コマンド | `ExecuteCDP` | `ExecuteBiDi` |
| 公式仕様 | [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/) | [WebDriver BiDi](https://w3c.github.io/webdriver-bidi/) |

## CDP を選ぶとき

- フォーム入力・クリック・XPath / CSS・Shadow / iframe が主戦場
- 既存の `Demo_CDP` やコミュニティ事例をそのまま応用したい
- DevTools のドメイン（Network, DOM, Page…）を直接叩く予定

## BiDi を選ぶとき

- `session.subscribe` ベースのイベント購読を標準に寄せたい
- ダイアログ制御や将来の W3C 互換を意識したい
- BiDi で足りない部分だけ CDP にフォールバックする構成にしたい

## 混ぜるパターン

### 1. BiDi → CDP コンテキスト

```vb
Dim bidiTab As WebDriverBiDiContext
Set bidiTab = ShSetting01_StartBrowser.StartBiDiModeContext("https://example.com")

Dim cdpTab As CDPContext
Set cdpTab = bidiTab.ConvertToCDPContext
cdpTab.getElementByQuery("button").click
```

### 2. BiDi+ で CDP コマンドを中継

```vb
' goog:cdp.getSession / goog:cdp.sendCommand
' 詳細は「低レイヤー BiDi / CDP コマンドについて」ガイドへ
```

→ [低レイヤー BiDi / CDP コマンドについて](/guides/extend-raw-protocol)

## おすすめの始め方

1. [はじめに](/getting-started) で CDP Hello World
2. 同じ操作を BiDi でも試す（コードグループ参照）
3. 要素操作が必要な画面は CDP（または `ConvertToCDPContext`）

迷ったら **CDP を主、BiDi をイベント／将来互換用** で問題ありません。
