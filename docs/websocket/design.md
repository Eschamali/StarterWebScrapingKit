---
description: Pipe 版に加えて WebSocket モードを増設した理由と設計。RFC 6455 ベースの拡張として Pipe ロジックとどう接続するかを説明します。
---

# 設計思想について

このツールは元々、`--remote-debugging-pipe` 1本で研ぎ澄ましてきました。
しかし、WebSocket モードだからこそできるいくつかの特有の機能があることがわかり、拡張ポジションとして増設しました。

「WebSocket 仕様書（[RFC 6455](https://datatracker.ietf.org/doc/html/rfc6455)）」を基にペイロードデータを取り出しつつ、ペイロードデータ終了の合図が来たらヌル文字を付与して Pipe 版ロジックに合わせる、といった感じで比較的簡単に増設できました。

通信方式こそ`WebSocket`ですが、一般的なWebSocketライブラリとは違う思想で作っています。  
当然ですが、VBAには汎用性のあるWebSocketライブラリは...いや実はあるっちゃあります。

**[wasabi](https://github.com/vbacollective/wasabi)**

TCP・WebSocket・TLS・MQTT まで1ファイルで扱える、とても多機能な汎用ライブラリです。  
実は私も、これで拡張できないか検討しました。が、**CDP 専用として使うには合わない点**がいくつかあり、実務としては採用できませんでした🙃

::: tip この節の読み方
「wasabi が悪い」という話ではありません。**汎用ライブラリと、CDP 専用の部品とでは、求められるものが違う**、という話です。wasabi 側の挙動は、`Wasabi.bas`（v2.3.8-beta）のソースを読んで確認した内容です。
:::

## 理由1：握手に付く `Origin` ヘッダーで、接続を拒否される

WebSocket の接続は、最初に HTTP で「握手」をします。このとき wasabi は、`Origin: http://<ホスト>:<ポート>` を**必ず付けて**送ります。付与は固定で、あとから追加できるカスタムヘッダーは「足す」だけなので、`Origin` を外すことはできません。

一方、CDP のデバッグサーバーには、次のようなセキュリティ仕様があります。

- 握手に `Origin` ヘッダーが含まれる場合（Web ページからの接続など）は、起動引数 `--remote-allow-origins=<origin>`（または `*`）に一致している必要がある。一致しなければ **403 Forbidden**
- `Origin` ヘッダーが無い接続（CLI ツール、Puppeteer、Node.js など）は、既定で許可される

つまり wasabi で CDP につなぐと、ブラウザ側に `--remote-allow-origins=*` を付けて起動しない限り、`HTTP/1.1 403 Forbidden` で拒否されます。私の環境でも、そのとおりになりました。

このツールの `CDPCoreViaWebSocket` は、`Host` / `Upgrade` / `Connection` / `Sec-WebSocket-Key` / `Sec-WebSocket-Version` だけの最小の握手を送ります。`Origin` が無いため、**起動引数の追加は不要**です。既に起動しているブラウザへの後付け接続（[`--remote-debugging-port` だけで起動したブラウザ](/websocket/capabilities)）も、そのままつながります。

## 理由2：メッセージキューが固定 512 件。あふれたら捨てられる

wasabi は、受信したメッセージを、接続ごとの **固定長のキュー（512 件）** に積みます。アプリ側が取り出しきれずに満杯になると、`message queue full, dropping message` という警告ログとともに、**そのメッセージは捨てられます**。

- バッファーサイズは `WebSocketSetBufferSize` で最大 16 MB まで広げられますが、**広がるのは受信用／断片用のバッファーで、キューの 512 件は変わりません**（定数のため、変えるには `Wasabi.bas` を書き換える必要があります）
- 1 メッセージが断片バッファーに収まらないほど大きい場合は、接続ごと切断されます

CDP は、ページを開くだけでも大量のイベント（`Network.*` など）が流れます。スクショ連発や大量の非同期実行をすると、キューに積まれる速さに取り出しが追いつきません。私の環境でも、バッファーを最大まで広げても `message queue full` が出続け、取りこぼしで処理が破綻しました🫠

Pipe 版には、そんな件数の上限はありません。`CDPCoreViaWebSocket` も同じで、**受信キューを持たず**、`CDPCore` が読んだ分をそのまま `ADODB.Stream` に蓄積していきます。メッセージ件数に上限はありません。

## 理由3：Pipe 版とフローを揃えられない

このツールの受信処理は、通信経路ごとに、取り出し口だけを差し替える設計です。Pipe 版の `TakePipe` と WebSocket 版の `TakeWebSocket`（どちらも `CDPCore` の内部処理）は、ほぼ同じ形をしています。

| 工程 | Pipe | WebSocket（`CDPCoreViaWebSocket`） |
| --- | --- | --- |
| 受信できる量の確認 | `PeekNamedPipe` | `ioctlsocket`（`FIONREAD`） |
| データの取り出し | `ReadFile` | `recv` |
| メッセージの終わり | 終端のヌル文字 | `FIN` フラグ。届いたらヌル文字を付け足して、Pipe 版と同じ形に揃える |
| 蓄積先 | `ADODB.Stream` | 同左 |
| 取り出しのループ | 終わりの合図が来て、かつ受信 0 になるまで | 同左 |
| エラー情報の持ち方 | `IsErrorStylePeekNamedPipe` / `IsErrorStyleReadFile` | 同名で Pipe に合わせている |

wasabi は、`WSAAsyncSelect` と隠しウィンドウのサブクラス化で受信を検知し、完成したメッセージをキューに積んで渡す、という別の設計です。そのため、上の流れに合わせるには `Wasabi.bas`（約 8,100 行の単一モジュール）の内部を読み解いて、改造する必要がありました。

「接続口だけ違って、中身のフローは同じ」という Pipe 版の見通しの良さを、WebSocket でも保ちたかったのです。

## 理由4：CDP に不要な機能が多く、API の使い方が目立つ

wasabi は汎用ライブラリなので、CDP では使わない機能がたくさん入っています。

| wasabi の機能 | CDP 制御での要否 |
| --- | --- |
| TLS/SSL（Schannel） | 不要。CDP は `ws://`（平文）で足ります |
| プロキシ（NTLM / Kerberos 対応） | 不要 |
| MQTT | 不要 |
| permessage-deflate（圧縮） | 不要。圧縮を要求しない握手で足ります |
| 自動再接続、オフラインキュー、RTT 計測、Happy Eyeballs | 不要 |
| 暗号 API による握手キーの生成と検証 | 不要。ローカルの CDP では固定キー、応答は `101` の確認だけで足ります |

加えて、wasabi は内部で、**実行可能なメモリを確保して、機械語を書き込んで呼び出す**処理（`VirtualAlloc` → `VirtualProtect`）や、隠しウィンドウのサブクラス化を行っています。高速化のための正当な実装ですが、セキュリティソフトが警戒しやすい API の使い方でもあります。業務で配布するマクロブックとしては、**誤検知のリスクをなるべく減らしたい**ところです。

一方、`CDPCoreViaWebSocket` の WinSock 関連の API 宣言は、`socket` / `connect` / `send` / `recv` / `ioctlsocket` など、10 個だけです。

::: warning ちなみに
このツールにも、WebView2 経路には機械語のサンクがあります。ここで比べているのは、**WebSocket 経路どうし**の話です。
:::

## まとめ：CDP 専用に割り切った

| 観点 | wasabi | CDPCoreViaWebSocket |
| --- | --- | --- |
| 握手 | `Origin` 固定付き | `Origin` なし（起動引数の追加が不要） |
| 受信 | 固定 512 件のキュー。満杯で破棄 | キューなし。`ADODB.Stream` に蓄積 |
| 受信のモデル | 非同期（隠しウィンドウ）＋キュー | Pipe 版と同じ、取り出しループ |
| 規模 | 約 8,100 行（TCP / TLS / MQTT など込み） | 約 1,400 行（CDP の接続探索込み） |
| 機械語・サブクラス化 | あり | なし |

> そんなのは嫌だ🥺WebSocket経由でも、Pipe版みたいなロジックに仕上げてやる🔥

こうして、独自でCDP専用として実装しました🤠  
WinSock で自作するまでの道のりは、[このツールの誕生秘話](/stories/birth-story)で語っています。

こういった拡張で実装しているため、Pipe 版での起動とは少し異なります。[次のページ](/websocket/capabilities)にて説明します。

## 関連

- [WebSocket モードでできること](/websocket/capabilities)
- [アーキテクチャ](/concepts/architecture)
- [このツールの誕生秘話](/stories/birth-story) — WinHTTP を諦めて WinSock を自作した話
- [Transport 層とバッファ管理](/core-comparison/transport) — `ws` パッケージに委譲する Node 勢との比較
