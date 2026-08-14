---
description: Pipe / WebSocket の Transport 層と、生バイトから JSON を1件ずつ切り出すバッファ管理を Puppeteer / Playwright と比較。NUL 区切りフレーミングと O(n²) 回避策を検証します。
---

# Transport 層とバッファ管理

Pipe も WebSocket も、結局は **1本の管に JSON がひたすら流れてくるだけ**です。届いた生バイトを貯め、文字列に戻し、1件分の JSON を切り出す ―― この一番地味な層から見ていきます。

## 1. どちらの管を本流にするか

3者とも Pipe と WebSocket の両方を実装していますが、**デフォルトがきれいに分かれます**。

| | Pipe 実装 | WebSocket 実装 | 自前起動時のデフォルト |
| --- | --- | --- | --- |
| **Playwright** | `server/pipeTransport.ts` | `server/transport.ts` | **Pipe** |
| **Puppeteer** | `node/PipeTransport.ts` | `node/NodeWebSocketTransport.ts` | **WebSocket** |
| **StarterWebScrapingKit** | `CDPCore.cls` | `CDPCoreViaWebSocket.cls` | **Pipe** |

Playwright は `defaultArgs` で `--remote-debugging-pipe` を無条件に付けるうえ、ユーザーが自分で同じ引数を渡そうとすると例外を投げます。

```ts
// playwright-core/src/server/chromium/chromium.ts
chromeArguments.push('--remote-debugging-pipe');
// ...
if (args.find(arg => arg.startsWith('--remote-debugging-pipe')))
  throw new Error('Playwright manages remote debugging connection itself.');
```

対して Puppeteer は `ChromeLauncher` の既定値が `pipe = false` で、明示的に `pipe: true` を渡さない限り WebSocket 経由になります。

```ts
// puppeteer-core/src/node/ChromeLauncher.ts
pipe = false,
```

::: tip 設計判断としての「Pipe 本流」
このキットが Pipe を本流に据えている（[アーキテクチャ](/concepts/architecture) 参照）のは、偶然にも **Playwright と同じ判断**です。ポートを開かないぶん外部からの割り込みを受けず、自分で起動したブラウザを確実に掌握できるためで、同一 PC での自動化ではこちらが素直です。

WebSocket は「すでにログイン済みのブラウザに後から乗る」という Pipe では原理的に不可能な用途のために増設されています（[WebSocket モードの設計思想](/websocket/design)）。
:::

## 2. Pipe のフレーミングは3者とも同じ

Pipe には WebSocket のようなフレーム構造がありません。そこで CDP は **`\0`（NUL 文字）区切り**という単純なルールでメッセージを区切ります。これはプロトコル側の仕様なので、3者とも従うしかない部分です。

```ts
// playwright-core/src/server/pipeTransport.ts
send(message: ProtocolRequest) {
  this._pipeWrite.write(JSON.stringify(message));
  this._pipeWrite.write('\0');
}
```

問題は受信側です。TCP やパイプは「1回の読み込み ＝ 1件のメッセージ」を保証しないため、

- 1回の読み込みに**複数件**まとまって届くこともあれば、
- 1件が**途中でぶつ切り**にされて複数回に分かれて届くこともある

という前提でバッファを組む必要があります。

## 3. 「毎回作り直さない」という共通の課題

素朴に実装すると、受信のたびにバッファ全体を連結し、1件取り出すたびに残り全体をコピーし直すことになります。これはメッセージ数に対して O(n²) に効いてくるため、**3者ともこれを避ける工夫を入れています**。ただしアプローチは言語ごとに違います。

### Node 側 ― 配列に貯めて、見つかった時だけ結合する

```ts
// playwright-core/src/server/pipeTransport.ts
_dispatch(buffer: Buffer) {
  let end = buffer.indexOf('\0');
  if (end === -1) {
    this._pendingBuffers.push(buffer);   // まだ揃わない → 貯めるだけ
    return;
  }
  this._pendingBuffers.push(buffer.slice(0, end));
  const message = Buffer.concat(this._pendingBuffers).toString();  // 揃った時だけ結合
  // ... 同一チャンク内の2件目以降を while で回収 ...
  this._pendingBuffers = [buffer.slice(start)];   // 端数は次回へ持ち越し
}
```

Puppeteer の `#dispatch` もほぼ同型です。要点は **`Buffer.concat` を「区切りが見つかった時」まで遅延させている**ことと、探索が `Buffer.indexOf`（ネイティブ実装）である点です。

### VBA 側 ― 最初から大きく確保して、その場で書き換える

VBA の文字列は不変（immutable）なので、`&` 連結も `Right()` による切り出しも毎回バッファ全体のコピーが走ります。そこで `CDPCore.cls` は真逆のアプローチを取っています。

```vb
' 1MB を初期確保。足りなくなったら倍々に拡張する（再確保自体を稀にする）
Private Const InitialBuffer As Long = 2 ^ 20

If .EndCursor + resSize > .length Then
    .strBuffer = .strBuffer & String$(.length, vbNullChar)
    .length = .length * 2
End If
```

```vb
' 連結ではなく、Mid$ ステートメントで「その場書き換え」（再確保が起きない）
Mid$(responseCDP.strBuffer, responseCDP.AddCursor + 1) = Utf8Converter.BytesToString((.Read))
```

```vb
' 取り出しはカーソルを進めるだけ。バッファ本体は書き換えない
EndPos = searchNull(.strBuffer, .StartCursor + 1)
TakeCDPMessage = Mid$(.strBuffer, .StartCursor + 1, EndPos - .StartCursor - 1)
.StartCursor = EndPos
```

NUL の探索も、1文字ずつ走査するのではなく VBA 処理系がネイティブ実装している `InStr` に任せています。

```vb
' CDP messages received from chrome are null-terminated
' Updated: 25/10/25: Daniel Polak - new faster version
lngPos = InStr(StartPos, checkString, vbNullChar, vbBinaryCompare)
```

コメントに改善履歴が残っている通り、ここは一度パフォーマンスの壁にぶつかってから意図的にチューニングされた箇所です。

### 対応関係の整理

| 段階 | Playwright / Puppeteer | StarterWebScrapingKit |
| --- | --- | --- |
| ① 生バイトを貯める | `pendingBuffers.push(buffer)` | `CDPPipeStream.Write res`（`ADODB.Stream`） |
| ② バイト列を文字列に戻す | `Buffer.concat(...).toString()` | `ReadyCDPMessage`（`Mid$` でバッファへ流し込み） |
| ③ 区切りで1件ずつ取り出す | `while (indexOf('\0') !== -1)` | `TakeCDPMessage`（`InStr` ＋ カーソル前進） |
| コピー回数を減らす工夫 | 結合を「揃った時だけ」に遅延 | 事前確保 ＋ 倍々拡張で再確保を稀にする |
| 探索関数 | `Buffer.indexOf`（ネイティブ） | `InStr`（ネイティブ） |

::: info 粒度の違い
Node 側は `.toString()` の一行が「デコード」と「境界探索」を兼ねていますが、VBA 側は `ReadyCDPMessage`（デコード）と `TakeCDPMessage`（切り出し）という**別々のプロシージャに分離**されています。層の切り方としてはむしろこちらのほうが細かい粒度です。
:::

### 探索するタイミングが逆になっている

地味ですが面白い違いがあります。

- **Node 側**：`buffer.indexOf('\0')` ―― まだ**バイトの世界**にいるうちに区切りを探し、見つかった分だけデコードする
- **VBA 側**：溜まっているバイトを先に丸ごと文字列化してから、**デコード後の文字列**に対して `InStr` で探す

順番が逆ですが、結果は変わりません。UTF-8 では NUL（`0x00`）がマルチバイト文字の一部として現れることが仕様上あり得ないため、どちらの世界で探しても境界は一致します。

## 4. WebSocket ―― ここだけは土俵が違う

Pipe の話とは打って変わって、WebSocket では**比較の前提そのものが崩れます**。

### Node 側は、そもそも書いていない

Playwright と Puppeteer のソースを検索しても、RFC 6455 のフレーム解析（FIN ビット、opcode、マスク処理、126 / 127 の拡張長分岐）は**一行も出てきません**。

```ts
// playwright-core/src/server/transport.ts
import ws from 'ws';
```

```ts
// puppeteer-core/src/node/NodeWebSocketTransport.ts
import NodeWebSocket from 'ws';
```

両者とも npm の `ws` パッケージに完全に委譲しており、`message` イベントが飛んできた時点で解析済みの完成メッセージが渡ってきます。

### このキットは、生 WinSock で自前実装している

一方 `CDPCoreViaWebSocket.cls` は `ws2_32.dll` を直接叩き、フレーム解析を手書きしています。クラス冒頭のコメントに、その判断理由がそのまま残っています。

> 「WinHttpWebSocket○○」の場合、`ioctlsocket` や `PeekNamedPipe` と言った「覗き見機能」はなく（中略）悩みに悩んだ結果、「WinSock」でほぼ1から作り上げました。これにより、Pipe 版と同じロジックとして運用することが可能になりました。

Windows 標準の `WinHttpWebSocketReceive` は「非同期コールバック」か「同期ブロッキング」の二択で、**「届いているか覗き見るだけ」ができません**。Pipe 版が `PeekNamedPipe` で実現している非ブロッキングのポーリング作法を WebSocket 側でも成立させるために、下回りごと作り直したというのが経緯です。

```vb
' ioctlsocket(FIONREAD) で受信済みバイト数だけ先に確認する（PeekNamedPipe と同じ役割）
Private Const FIONREAD As Long = &H4004667F
retIoctl = ioctlsocket(m_hSocket, FIONREAD, bytesAvailable)
```

```vb
' 送信フレームの組み立て。FIN ビット＋opcode を自分で立てる
frame(0) = &H80 Or opcode
```

```vb
' 受信ヘッダーから opcode を取り出し、126 / 127 の拡張長分岐も手書き
.opcode = (header(0) And &HF)
```

クライアントからサーバーへの送信は仕様上マスク必須なので、4バイトのランダムマスクキー生成と XOR 処理も自前です。

### 「同じ発想」だが「同じ完成度」ではない

::: warning 現状の実装上の割り切り
フレーム解析の核心部分は動いていますが、`ws` パッケージが長年かけて作り込んできた**周辺仕様は意図的に省略されています**。

- **Ping / Pong 応答**（opcode `0x9` / `0xA`）の自動処理は未実装
- **ハンドシェイク検証**は `InStr(response, "HTTP/1.1 101") > 0` というステータス行の確認のみで、`Sec-WebSocket-Accept` ヘッダー（送った `Sec-WebSocket-Key` を SHA-1 ＋ Base64 したもの）の照合は行っていない
- `Sec-WebSocket-Key` 自体も RFC 6455 の例文値を固定で送っている
- `permessage-deflate` などの拡張ネゴシエーションは非対応

自分で起動したローカルの Chromium に繋ぐという用途に限れば実害が出にくい範囲ですが、汎用 WebSocket クライアントとして任意のサーバーに繋ぐ用途は想定していません。
:::

つまりこの層は「実装力で並んだ / 並ばない」という話ではなく、**JS 開発者は書く必要すらなかったコードを、VBA にはそれに相当するライブラリが存在しないため自力で書いた**、という構図です。

## まとめ

| | Pipe のバッファ管理 | WebSocket のフレーム解析 |
| --- | --- | --- |
| **Playwright / Puppeteer** | チャンク配列 ＋ 遅延 `concat` ＋ `indexOf` | `ws` ライブラリに委譲（自前コードなし） |
| **StarterWebScrapingKit** | 事前確保 ＋ 倍々拡張 ＋ `Mid$` 書換 ＋ `InStr` | 生 WinSock で RFC 6455 を自前実装 |

- **Pipe** は本当に対等です。言語の得意技（`Buffer` の配列操作 / 文字列の `Mid$` 直接書き込み）に合わせて、それぞれ別ルートから「コピー回数を最小化する」という同じ最適解に到達しています。
- **WebSocket** は比較になりません。Node 側は業界標準ライブラリに丸投げできるので、そもそも競技に参加していないためです。

::: tip この層の検証について
バッファ管理は `Test_AsyncBenchmark.bas` で実負荷にかけられています。30 タブ × 10 ラウンドで最大 300 件の Base64 スクリーンショットを流し込み、Pipe / WebSocket の両経路で完走するかを見るものです（[テストの現状](/core-comparison/gaps)）。
:::

## 次に読む

- [ディスパッチとイベント配信](/core-comparison/dispatch) — 切り出した JSON をどこへ届けるか
- [WebSocket モードの設計思想](/websocket/design) — 増設の経緯
- [アーキテクチャ](/concepts/architecture) — `CDPCore` / `CDPCoreViaWebSocket` の位置づけ
