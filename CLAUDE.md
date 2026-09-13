# 令和のIE (Edge-CDP-VBA)プロジェクト

## 1. プロジェクト概要
外部ドライバ（msedgedriver.exe）やSeleniumBasic、管理者権限、外部DLLのインストールを一切行わず、  
Excel VBA（標準モジュール/クラスモジュール）とWindows標準APIのみを用いて、  
Microsoft Edge(Chromium系)をChrome DevTools Protocol (CDP) 経由で自動操作するスタンドアロン型ライブラリを構築する。  
開発言語は日本語。応答も日本語で行う。

## 2. 基本アーキテクチャ
- **通信レイヤー**
  - `--remote-debugging-pipe`による匿名パイプ通信
  - `--remote-debugging-port`によるWebSocket(WinSock)通信
  - WebView2の`CallDevToolsProtocolMethod(ForSession)`によるCOM通信

- **プロトコル**: Chrome DevTools Protocol (CDP) の JSON-RPC 2.0
- **文字コード**: VBA内部（UTF-16LE）と CDP（UTF-8）の相互変換を徹底

## 3. 【最重要】絶対禁止事項（Constraints）
1. `SeleniumBasic` や外部COM DLLの登録を前提としたコードは一切書かないこと
2. `msedgedriver.exe` や `chromedriver.exe` などの外部バイナリを配置・要求しないこと
3. 管理者権限が必要なレジストリ操作やインストーラー実行を行わないこと
4. VBAの代わりにPython、PowerShellスクリプト等を別ファイルとして生成・実行させないこと（VBA単体完結が目標）
5. GitのCommit,Push等といったGitに変更を加える操作を禁止。ここは人間が精査するところです
6. 機能を増やす依頼は基本的に、Extensionsブランチの`Extensions/`配下として作成すること。`src/`配下編集による機能追加/改修等はユーザーからの明示的な指示がない限り、しない

## 4. VBAコーディング規約
- 行継続記号（アンダースコア）は、25回が上限。また、その記号の右側に同じ行内でコメントを書くと構文エラー
- モジュールレベル変数・定数はモジュール冒頭（全プロシージャより手前）に置く
- Extensionsブランチでのみ出現する`Extensions/`配下のVBAsourceは、`ShiftJis`として保存すること
- `BiDi+`を扱う場面(goog:cdp.sendCommand等)が必要な場合はまず、`WebDriverBiDiContext.UpgradeBiDiPlus`で賄えるか確認する。賄えない場合は理由を添えてユーザーに判断を委ねること

## 5. 主要なファイル構成と「真実の所在」

| 場所 | 中身 | 性質 |
|---|---|---|
| `src/` | VBEでエクスポートする際のsource一式。 | **★正★ ここを編集する** |
| `ForDevelopers\OperationCheck/` | テストコード一式。 | CDPとBiDi用に基本分けている |
| `ForDevelopers\TemplateExtensions` | 機能拡張用のテンプレート。 | CDPとBiDi用に基本分けている |
| `ForAI/` | 他のWeb自動化ツールのsource一式 | アイデア出し用。.gitignoreに登録済み |
| `assset/` | ChromiumでWebDriverBiDi化するやつ | WebDriverBiDiが動くのはこれのおかげ |
| `docs/` | VitePress製ドキュメント | gh-pagesブランチでのみ出現。 |
| `Extensions/` | このツールの拡張機能一式(DL監視等) | Extensionsブランチでのみ出現。 |
