---
description: Pipe 版に加えて WebSocket モードを増設した理由と設計。RFC 6455 ベースの拡張として Pipe ロジックとどう接続するかを説明します。
---

# 設計思想について

このツールは元々、`--remote-debugging-pipe` 1本で研ぎ澄ましてきました。
しかし、WebSocket モードだからこそできるいくつかの特有の機能があることがわかり、拡張ポジションとして増設しました。

「WebSocket 仕様書（[RFC 6455](https://datatracker.ietf.org/doc/html/rfc6455)）」を基にペイロードデータを取り出しつつ、ペイロードデータ終了の合図が来たらヌル文字を付与して Pipe 版ロジックに合わせる、といった感じで比較的簡単に増設できました。

こういった拡張で実装しているため、Pipe 版での起動とは少し異なります。[次のページ](/websocket/capabilities)にて説明します。

## 関連

- [WebSocket モードでできること](/websocket/capabilities)
- [アーキテクチャ](/concepts/architecture)
- [Transport 層とバッファ管理](/core-comparison/transport) — `ws` パッケージに委譲する Node 勢との比較
