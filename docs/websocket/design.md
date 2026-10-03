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

実は私も最初はこれで拡張しようと考えていました。が、以下の問題が発生してしまい、実務としては採用できませんでした🙃
- 巨大レスポンスに弱い
  - スクショ連発や大量の非同期実行を行うと、すぐに破綻しました🙃
- セキュリティ誤検知しやすい構造
- CDPと関係のない機能がある

特に1つ目は致命的です。Webスクレイピングでは普通のことなのに、wasabiにやらせるとすぐにエラー🫠  
Pipe版ではそんなレスポンス制限なんて気にしなくていいのに、WebSocket経由はそういう意識をしないといけない...

> そんなのは嫌だ🥺WebSocket経由でも、Pipe版みたいなロジックに仕上げてやる🔥

こうして、独自でCDP専用として実装しました🤠

こういった拡張で実装しているため、Pipe 版での起動とは少し異なります。[次のページ](/websocket/capabilities)にて説明します。

## 関連

- [WebSocket モードでできること](/websocket/capabilities)
- [アーキテクチャ](/concepts/architecture)
- [Transport 層とバッファ管理](/core-comparison/transport) — `ws` パッケージに委譲する Node 勢との比較
