# 令和のIE (Edge-CDP-VBA)プロジェクト

## 1. プロジェクト概要
外部ドライバ（msedgedriver.exe）やSeleniumBasic、管理者権限、外部DLLのインストールを一切行わず、  
Excel VBA（標準モジュール/クラスモジュール）とWindows標準APIのみを用いて、  
Microsoft Edge(Chromium系)をChrome DevTools Protocol (CDP) 経由で自動操作するスタンドアロン型ライブラリを構築する。  
開発言語は日本語。応答も日本語で行う。

## 2. 基本アーキテクチャ
- **ブラウザ起動**: WindowsAPIの `CreateProcess` から `--remote-debugging-pipe` でEdgeを起動
- **通信レイヤー**: Windows標準API（`CreatePipe` または `Winsock`）をVBAから直叩き
- **プロトコル**: Chrome DevTools Protocol (CDP) の JSON-RPC 2.0
- **文字コード**: VBA内部（UTF-16LE）と CDP（UTF-8）の相互変換を徹底

## 3. 【最重要】絶対禁止事項（Constraints）
1. `SeleniumBasic` や外部COM DLLの登録を前提としたコードは一切書かないこと
2. `msedgedriver.exe` や `chromedriver.exe` などの外部バイナリを配置・要求しないこと
3. 管理者権限が必要なレジストリ操作やインストーラー実行を行わないこと
4. VBAの代わりにPython、PowerShellスクリプト等を別ファイルとして生成・実行させないこと（VBA単体完結が目標）
5. GitのCommit,Pushを禁止。ここは人間が精査するところです

## 4. VBAコーディング規約
- 行継続記号（アンダースコア）は、25回が上限。また、その記号の右側に同じ行内でコメントを書くと構文エラー
- モジュールレベル変数・定数はモジュール冒頭（全プロシージャより手前）に置く

## 5. 主要なファイル構成と「真実の所在」

| 場所 | 中身 | 性質 |
|---|---|---|
| `src/` | VBEでエクスポートする際のsource一式。 | **★正★ ここを編集する** |
| `ForDevelopers\OperationCheck/` | テストコード一式。 | CDPとBiDi用に基本分けている |
| `ForDevelopers\TemplateExtensions` | 機能拡張用のテンプレート。 | CDPとBiDi用に基本分けている |
| `ForAI/` | 他のWeb自動化ツールのsource一式 | アイデア出し用。.gitignoreに登録済み |
| `assset/` | ChromiumでWebDriverBiDi化するやつ | WebDriverBiDiが動くのはこれのおかげ |
| `docs/` | VitePress製ドキュメント | gh-pagesブランチでのみ出現。 |
