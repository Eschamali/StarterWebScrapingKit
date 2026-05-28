# WebSocketViaAnonymousPipe - 匿名パイプ経由の CDP WebSocket 中継モジュール

## 1. 概要
本モジュールは、Excel VBA から Chromium 系ブラウザ（Chrome、Edge 等）の CDP（Chrome DevTools Protocol）を制御するための **セキュリティ・バイパス型 WebSocket 中継器 (Bridge)** です。

Windows API を用いた双方向の非ブロッキング・パイプラインを介して Excel と PowerShell を接続し、さらに PowerShell から WebSocket を通じて Chromium と通信を行うことで、安全かつ堅牢なブラウザ自動制御を実現します。

### 構成イメージ
```
Excel VBA  ⇄ [匿名パイプ (Anonymous Pipe)] ⇄  PowerShell (Bridge)  ⇄ [WebSocket] ⇄  Chromium
```

---

## 2. 背景と解決策 (セキュリティ・バイパス)

### 【背景】現代のセキュリティソフトによる制限
通常、VBA から `Shell` 関数などを用いて PowerShell を起動する際、引数（`-Command` など）に実行したいスクリプトのコードを直接渡します。しかし、これは現代の高度なセキュリティソフト（Norton、CrowdStrike、Windows Defender 等）が監視する「**Command Line Inspection (コマンドライン検査)**」の格好の標的となり、危険な処理として不当にブロックされるケースが多発します。

### 【解決策】標準入力 (StdIn) を用いたコード動的注入
この問題を回避するため、本モジュールは以下の仕組みを採用しています。
1. **「手ぶら」起動**: PowerShell プロセス（`powershell.exe`）を、引数を一切持たないクリーンな状態で起動し、セキュリティ検閲をすり抜けます。
2. **動的注入 (StdIn)**: プロセスが正常に開通したことを確認した後、API 経由で「標準入力（StdIn）」ストリームから PowerShell スクリプトコードを動的に流し込みます。

これは、ユーザーが Terminal（黒い画面）から手動でコマンドをタイピングして実行する「正当な対話操作」をプログラム上で完全に再現するものであり、セキュリティ検知を効果的に回避しながら本来の実行能力を安全に解放します。

---

## 3. 主な役割とニッチな活用例
本モジュールは本来、`remote-debugging-pipe` 特化で構築されているコアシステムに対して、堅牢性を損なわずに `remote-debugging-port` との直接接続を実現するための「拡張機能」として用意されました。
セルに所定の PowerShell コードを埋め込んで起動しておくことで、以下のような高度で特殊な自動制御が可能になります。

- **Android 実機の Chromium 制御**: `chrome://inspect` などを経由した任意のデバイス内のブラウザ制御。
- **WebView2 アプリケーションの制御**: 環境変数 `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` に `--remote-debugging-port=9222` を付与し、埋め込み WebView2 を外部から自動制御。
- **既存ブラウザへのアタッチ**: デバッグ起動している目の前のブラウザの特定のタブや本体にアタッチした自動制御。

---

## 4. 構成ファイル

| ファイル名 | 役割 |
| :--- | :--- |
| **`StartConnectWebSocketForChromium.ps1`** | **PowerShell 側の中継器 (Bridge) 実体**。<br>VBA から渡された環境変数ハンドルを基に匿名パイプを作成し、Chromium の WebSocket と接続して非同期に双方向データ中継を行います。 |
| **`PowerShellViaStdPipe.cls`** | **VBA 側のパイプ・プロセス制御クラス**。<br>Windows API (`CreateProcess`, `CreatePipe`) を駆使して PowerShell プロセスを非表示（`CREATE_NO_WINDOW`）で起動し、環境変数の管理や CDP 用の匿名パイプハンドル作成を行います。 |
| **`Demo_WebSocketViaAnonymousPipe.bas`** | **モジュールの利用デモ・設定用 VBA 標準モジュール**。<br>セルから読み込んだ中継器スクリプトの起動設定 (`AutoSetup`) や、WebView2 用のデバッグポート切り替え、ブラウザへの再アタッチ手順が定義されています。 |

---

## 5. 通信制御の仕組み (ハイブリッド判定ロジック)
VBA からの JSON コマンド送信および Chromium からのレスポンス受信は、匿名パイプのデータパケットがネットワークや OS のバッファによって細切れに分割されるリスクを考慮し、**ハイブリッド判定ロジック**によって制御されています。

### VBA (匿名パイプ) ➡️ Chromium (WebSocket)
- **高速ルート：直通便 🚀**
  - 受信したパケットが短文で、末尾がヌル文字（`0x00`）で終わっている場合、バッファに蓄積せず即座に Chromium へ転送します。
- **蓄積ルート：慎重便 📦**
  - パケットが分割されている（末尾がヌル文字でない）場合、ヌル文字が届くまで一時的に `MemoryStream` バッファにデータを蓄積し、完全にガッチャンコされた（合体した）1つの巨大な塊にしてから Chromium へ転送します。

### Chromium (WebSocket) ➡️ VBA (匿名パイプ)
- 受信した WebSocket のメッセージが完全に終了したタイミング（`EndOfMessage`）で、VBA 側が待機している終了マーカーである**ヌル文字（`0x00`）**を自動で追記してパイプを `Flush` 送信します。これにより、VBA 側でメッセージの終端を極めて正確に検知できます。

---

## 6. 主要プロシージャと設定 (VBA 側)

### `AutoSetup`
Excel のセル（既定では `A1`）に手動で貼り付けられた `StartConnectWebSocketForChromium.ps1` を読み込み、環境変数で接続パラメータ（ポート `9222` や CDP 入出力ハンドルなど）を継承させた上で、PowerShell プロセスを起動・初期化します。
- **セキュリティ配慮**: 初回利用時は、セキュリティの観点から使用者自身で手動でセルにスクリプトコードを配置する必要があります。
- **表示設定**: `ShowConsoleWindow = True` に設定すると、中継時の標準出力やログ（ミリ秒付き）を黒いコンソール画面で確認しながらデバッグできます（本番時は `False` で完全に非表示にできます）。

### `WebView2のクイックデバッグ切り替え`
環境変数 `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` に対して `--remote-debugging-port=9222` を動的に設定・削除します。これにより、VBA から同一プロセスで起動した WebView2 のデバッグポートを自在に開閉できます。

### `WebSocketによる冒険の始まり`
再接続処理（`reattach`）を用いたブラウザ操作のテンプレートです。
1. 既存の `targetID`（操作中のタブ）に再接続を試みます。
2. もしブラウザ側でタブが閉じられて接続を失っていた場合は、自動的に新しいタブの取得（`getTab`）または新規作成（`newTab`）を行って制御を復帰させます。

---

## 7. 導入・利用手順
1. **スクリプトのセル配置**: セキュリティ警告や誤検知を防ぐため、`StartConnectWebSocketForChromium.ps1` の全コードを Excel の適当なセル（例: `Sheet1` の `A1`）に手動で貼り付けます。
2. **対象ブラウザのデバッグ起動**: 操作対象のブラウザをリモートデバッグポートを有効にして起動します。
   - 例: `chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\tmp\chrome_profile"`
3. **初期セットアップ実行**: `Demo_WebSocketViaAnonymousPipe.AutoSetup` を実行します。裏で PowerShell 中継器が起動し、VBA とブラウザ間のブリッジが開通します。
4. **自動制御コードの実行**: `CDPBrowser` 等のインスタンスから `reattach` を実行し、ブラウザ操作（ページ遷移や要素抽出など）を開始します。
