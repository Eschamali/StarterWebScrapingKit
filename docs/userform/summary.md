---
description: UserForm へのモダンブラウザ埋め込み 3 手法（Edge / PowerShell / Excel 単体）の比較総括。実務での選び方の指針をまとめます。
---

# UserFormへのモダンブラウザ埋め込み：3つの手法の総括

本セクションでは、Excel VBAのUserFormにモダンなブラウザ（WebView2/Edge）を配置する3つのアプローチを紹介しました。  
最後に、それぞれの特性を比較し、実務でどの方針を採用すべきかの指針をまとめます。

::: info v3.0.0での更新
Lv.99（Excel単体）は、v3.0.0で本キットにネイティブ実装（`CDPCoreViaWebView2.cls`）されました。実装の複雑さは**もう気にする必要がありません**——利用者は `WebView2Form.StartCDPModeWebView2` を呼ぶだけです。これに伴い、Lv.1（Edge埋め込み）は廃止されています。以下の比較表・推奨ガイドはこの前提で更新しています。
:::

## 比較表

| 項目 | Lv.10 PowerShell 経由 | Lv.99 Excel 単体（v3.0.0〜標準実装） |
| --- | --- | --- |
| **実務での堅牢性** | ◎ (非常に安定) | 〇 (高度だが繊細、デバッグ中のブレークに注意) |
| **追加インストール** | 不要 (標準DLL流用) | 不要 (標準DLL流用) |
| **配布の容易さ** | 〇 (ps1同梱または動的生成) | ◎ (xlsmファイルのみ) |
| **フォーカス制御** | ◎ (ネイティブ同等) | ◎ (ネイティブ同等) |
| **見た目** | ◎ (真のWebView2) | ◎ (真のWebView2) |
| **プロセス構成** | Excel + PowerShell（別プロセス） | Excel単体（Excel.exe配下にWebView2） |
| **利用者から見た実装コスト** | 中 (名前付きパイプの往復を意識) | 低 (`StartCDPModeWebView2` を呼ぶだけ) |

## ユースケース別 推奨ガイド

::: tip 特にこだわりがなければ → Lv.99 Excel 単体
v3.0.0以降は、これが既定の選択肢です。外部プロセスを介さずExcel単体で完結し、利用者側のコードも `StartCDPModeWebView2` を呼ぶだけ。以前は「VBAの限界に挑戦するプロフェッショナル向け」でしたが、複雑さは本キット側に吸収済みです。

→ [Excel 単体の詳細](./vba-only)
:::

::: tip プロセスを分離しておきたい場合 → Lv.10 PowerShell 経由
「WebView2まわりの処理をExcel本体のプロセスから切り離しておきたい」「機械語サンクによるVBEブレーク時のクラッシュリスクを避けたい」といった事情があるなら、こちらも依然として有効な選択肢です。PowerShellが面倒なCOM制御を肩代わりしてくれます。

→ [PowerShell 経由の詳細](./powershell)
:::

::: warning 廃止：Lv.1 Edge 埋め込み
`msedge.exe` のKioskモードをUserFormにドッキングする手法は、v3.0.0でのネイティブWebView2実装に伴い廃止されました。当時の設計は[記録として](./edge)残しています。
:::

## おわりに

以前は「VBAにはIE（WebBrowserコントロール）しかない」と絶望視されていました。  
しかし、現在ではこれらの手法により、VBAでも最新のWeb技術（React, Vue, SPA, etc.）を駆使したリッチなUIを構築することが可能です。

プロジェクトの要件や環境に合わせて最適な手法を選び、Excel VBAの枠を超えた体験を是非作り上げてください！

Happy Coding with Modern Web in VBA!

## 関連

- [はじめに（UserFormコーナー）](./intro)
- [ページ遷移（CDP/BiDi）](/guides/navigation)
- [概要](/intro)
